"""Editor fotografico locale per le immagini fronte e retro dei biglietti."""
from dataclasses import dataclass
from pathlib import Path
import copy
import os
import queue
import tempfile
import threading
import tkinter as tk
from tkinter import colorchooser, filedialog, messagebox, ttk

from PIL import Image, ImageChops, ImageEnhance, ImageOps, ImageTk


RIEMPI = 'Riempi il riquadro (ritaglia)'
CONTIENI = 'Mostra tutta la foto (con bordi)'
ALLUNGA = 'Allunga al riquadro (deforma)'
EFFETTI = ('Nessuno', 'Bianco e nero', 'Seppia', 'Toni caldi', 'Toni freddi', 'Vintage', 'Negativo')
FORMATI = {'PNG': ('.png', 'PNG'), 'JPEG': ('.jpg', 'JPEG'), 'WebP': ('.webp', 'WEBP'), 'BMP': ('.bmp', 'BMP')}


@dataclass
class FotoModificata:
    originale: Image.Image
    impostazioni: dict
    immagine: Image.Image
    nome: str
    percorso_originale: str | None
    scala_x: float
    scala_y: float


def immagine_rgb(source):
    """Rispetta l'orientamento EXIF e appoggia la trasparenza su bianco."""
    picture = ImageOps.exif_transpose(source)
    if 'A' in picture.getbands() or 'transparency' in picture.info:
        rgba = picture.convert('RGBA')
        background = Image.new('RGBA', rgba.size, 'white')
        background.alpha_composite(rgba)
        return background.convert('RGB')
    return picture.convert('RGB')


def apri_immagine(percorso):
    with Image.open(percorso) as source:
        if source.width * source.height > 80_000_000:
            raise ValueError('La foto supera 80 megapixel: riducila prima di aprirla.')
        image = immagine_rgb(source)
        image.load()
        return image


def parametri_iniziali(raccomandate):
    return dict(modalita=RIEMPI, larghezza=raccomandate[0], altezza=raccomandate[1],
                zoom=100, posizione_x=50, posizione_y=50, rotazione=0,
                specchia=False, capovolgi=False, sfondo='#ffffff',
                luminosita=100, contrasto=100, saturazione=100, nitidezza=100,
                temperatura=0, tonalita=0, vignettatura=0, effetto='Nessuno',
                formato='PNG', qualita=95, dpi=300, proporzioni=True)


def validate_settings(s):
    w, h = s['larghezza'], s['altezza']
    if not isinstance(w, int) or not isinstance(h, int) or not (32 <= w <= 8000 and 32 <= h <= 8000) or w*h > 32_000_000:
        raise ValueError('Dimensioni consentite: da 32 a 8000 pixel per lato, massimo 32 megapixel.')
    if s['modalita'] not in (RIEMPI, CONTIENI, ALLUNGA) or s['effetto'] not in EFFETTI or s['formato'] not in FORMATI:
        raise ValueError('Seleziona una modalità e un formato validi.')
    if s['rotazione'] not in (0, 90, 180, 270):
        raise ValueError('Seleziona una rotazione valida.')
    if not 100 <= s['zoom'] <= 300 or not 0 <= s['posizione_x'] <= 100 or not 0 <= s['posizione_y'] <= 100:
        raise ValueError('Controlla zoom e posizione.')
    for key in ('luminosita', 'contrasto', 'saturazione', 'nitidezza'):
        if not 0 <= s[key] <= 200:
            raise ValueError('Controlla le regolazioni colore.')
    if not (-100 <= s['temperatura'] <= 100 and -180 <= s['tonalita'] <= 180 and 0 <= s['vignettatura'] <= 100):
        raise ValueError('Controlla tonalità, temperatura e vignettatura.')
    if not (10 <= s['qualita'] <= 100 and 72 <= s['dpi'] <= 1200):
        raise ValueError('Qualità fra 10 e 100; risoluzione di esportazione fra 72 e 1200 DPI.')
    try:
        from PIL import ImageColor
        ImageColor.getrgb(s['sfondo'])
    except (ValueError, TypeError):
        raise ValueError('Colore di sfondo non valido.')


def _orient(source, settings):
    image = source.copy()
    turn = settings['rotazione']
    transposes = {90: Image.Transpose.ROTATE_270, 180: Image.Transpose.ROTATE_180, 270: Image.Transpose.ROTATE_90}
    if turn:
        image = image.transpose(transposes[turn])
    if settings['specchia']:
        image = ImageOps.mirror(image)
    if settings['capovolgi']:
        image = ImageOps.flip(image)
    return image


def _layout(source, settings, w, h):
    """Restituisce anche i pixel originali per ciascun pixel di output."""
    sw, sh = source.size
    zoom = settings['zoom']/100.0
    cx, cy = settings['posizione_x']/100.0, settings['posizione_y']/100.0
    mode = settings['modalita']
    if mode == CONTIENI:
        scale = min(w/sw, h/sh)*zoom
        offset_x, offset_y = (w-sw*scale)*cx, (h-sh*scale)*cy
        coeff = (1/scale, 0, -offset_x/scale, 0, 1/scale, -offset_y/scale)
        image = source.transform((w,h), Image.Transform.AFFINE, coeff, resample=Image.Resampling.BICUBIC, fillcolor=settings['sfondo'])
        return image, 1/scale, 1/scale
    if mode == RIEMPI:
        ratio = w/h
        crop_w = min(sw, sh*ratio)/zoom
        crop_h = crop_w/ratio
    else:
        crop_w, crop_h = sw/zoom, sh/zoom
    left, top = (sw-crop_w)*cx, (sh-crop_h)*cy
    box = (left, top, left+crop_w, top+crop_h)
    image = source.resize((w,h), Image.Resampling.LANCZOS, box=box)
    return image, crop_w/w, crop_h/h


def _colors(image, settings):
    effect = settings['effetto']
    if effect == 'Bianco e nero':
        image = ImageOps.grayscale(image).convert('RGB')
    elif effect == 'Seppia':
        image = ImageOps.colorize(ImageOps.grayscale(image), '#2c1d11', '#ffe2a8').convert('RGB')
    elif effect == 'Vintage':
        image = ImageEnhance.Color(image).enhance(.55)
        image = ImageEnhance.Contrast(image).enhance(.87)
        image = Image.blend(image, Image.new('RGB', image.size, '#d2ac7e'), .14)
    elif effect == 'Negativo':
        image = ImageOps.invert(image)
    temperature = settings['temperatura'] + (30 if effect == 'Toni caldi' else -30 if effect == 'Toni freddi' else 0)
    if temperature:
        t = max(-130, min(130, temperature))/130.0
        red, green, blue = image.split()
        red = red.point([max(0,min(255,round(i*(1+.25*t)))) for i in range(256)])
        blue = blue.point([max(0,min(255,round(i*(1-.25*t)))) for i in range(256)])
        image = Image.merge('RGB', (red,green,blue))
    hue = settings['tonalita']
    if hue:
        hh, ss, vv = image.convert('HSV').split()
        delta = round(hue*255/360)
        hh = hh.point([(i+delta)%256 for i in range(256)])
        image = Image.merge('HSV', (hh,ss,vv)).convert('RGB')
    for key, enhancer in (('luminosita', ImageEnhance.Brightness), ('contrasto', ImageEnhance.Contrast),
                          ('saturazione', ImageEnhance.Color), ('nitidezza', ImageEnhance.Sharpness)):
        if settings[key] != 100:
            image = enhancer(image).enhance(settings[key]/100.0)
    if settings['vignettatura']:
        radial = Image.radial_gradient('L').resize(image.size, Image.Resampling.BILINEAR)
        strength = settings['vignettatura'] / 100 * .8
        mask = radial.point([round(i*strength) for i in range(256)])
        image = Image.composite(Image.new('RGB', image.size, 'black'), image, mask)
    return image


def elabora(originale, settings, preview_size=None):
    validate_settings(settings)
    w, h = settings['larghezza'], settings['altezza']
    if preview_size is not None:
        limit_w, limit_h = preview_size
        ratio = min(limit_w/w, limit_h/h, 1.0)
        w, h = max(1,round(w*ratio)), max(1,round(h*ratio))
    oriented = _orient(originale, settings)
    fitted, sx, sy = _layout(oriented, settings, w, h)
    return _colors(fitted, settings), sx, sy


def stima_dpi(originale, settings, box_ratio, inches):
    """DPI stimati dal campionamento reale, senza contare l'ingrandimento come dettaglio."""
    w,h = settings['larghezza'], settings['altezza']
    sw,sh = originale.size if settings['rotazione'] in (0,180) else originale.size[::-1]
    zoom = settings['zoom']/100
    if settings['modalita'] == RIEMPI:
        crop_w = min(sw, sh*w/h)/zoom
        crop_h = crop_w*h/w
        sx,sy = crop_w/w, crop_h/h
    elif settings['modalita'] == CONTIENI:
        scale = min(w/sw,h/sh)*zoom
        sx=sy=1/scale
    else:
        sx,sy=sw/zoom/w,sh/zoom/h
    visible_w=min(w,h*box_ratio)
    visible_h=min(h,w/box_ratio)
    return min(visible_w*min(1,sx)/inches[0], visible_h*min(1,sy)/inches[1])


def salva_immagine(path, image, settings):
    """Scrive una copia completa prima di sostituire il file scelto."""
    ext, image_format = FORMATI[settings['formato']]
    allowed = {'.jpg','.jpeg'} if image_format == 'JPEG' else {ext}
    if Path(path).suffix.lower() not in allowed:
        raise ValueError('L’estensione del nome deve corrispondere al formato selezionato: '+', '.join(sorted(allowed)))
    handle, temporary = tempfile.mkstemp(prefix='.foto_', suffix=ext, dir=str(Path(path).resolve().parent))
    os.close(handle)
    try:
        options = {'dpi':(settings['dpi'], settings['dpi'])}
        if image_format in ('JPEG','WEBP'):
            options['quality'] = settings['qualita']
        if image_format == 'WEBP':
            # WebP stores resolution in EXIF; Pillow's generic dpi option is ignored.
            from PIL.TiffImagePlugin import IFDRational
            exif=Image.Exif()
            exif[282]=IFDRational(settings['dpi'],1)
            exif[283]=IFDRational(settings['dpi'],1)
            exif[296]=2
            options['exif']=exif.tobytes()
        image.save(temporary, format=image_format, **options)
        os.replace(temporary,path)
    finally:
        if os.path.exists(temporary):
            os.unlink(temporary)


class EditorFoto:
    def __init__(self, parent, originale, nome, percorso, settings, box_ratio, inches, lato):
        self.parent = parent
        self.originale = originale
        self.preview_source = originale.copy()
        self.preview_source.thumbnail((1800, 1800), Image.Resampling.LANCZOS)
        self.nome = nome
        self.percorso = percorso
        self.original_settings = copy.deepcopy(settings)
        self.box_ratio = box_ratio
        self.inches = inches
        self.result = None
        self.updating = False
        self.pending = None
        self.preview_image = None
        self.window = tk.Toplevel(parent)
        self.window.withdraw()
        self.window.title('Modifica foto — '+lato)
        self.window.transient(parent)
        self.window.minsize(900, 630)
        self.window.geometry('1180x760')
        self.window.protocol('WM_DELETE_WINDOW', self.close)
        self.variables = {}
        for key,value in settings.items():
            cls = tk.BooleanVar if isinstance(value,bool) else tk.StringVar if isinstance(value,str) else tk.IntVar
            self.variables[key] = cls(value=value)
        self.comparison = tk.BooleanVar(value=False)
        self.status = tk.StringVar()
        self.preview_info = tk.StringVar()
        self.size_info = tk.StringVar()
        self._build(lato)
        for key,var in self.variables.items():
            var.trace_add('write', lambda *args, key=key: self.changed(key))
        self.comparison.trace_add('write', self.changed)
        self.window.update_idletasks()
        sw,sh=self.window.winfo_screenwidth(),self.window.winfo_screenheight()
        ww,wh=min(1180,max(900,sw-30)),min(760,max(630,sh-65))
        x=max(0,min(parent.winfo_rootx()+(parent.winfo_width()-ww)//2,sw-ww))
        y=max(0,min(parent.winfo_rooty()+(parent.winfo_height()-wh)//2,sh-wh))
        self.window.geometry(f'{ww}x{wh}+{x}+{y}')
        self.window.deiconify()
        self.window.grab_set()
        self.window.focus_set()
        self.schedule()

    def close(self):
        if self.pending:
            self.window.after_cancel(self.pending)
            self.pending=None
        self.window.destroy()
        if self.parent.winfo_exists():
            self.parent.grab_set()
            self.parent.focus_set()

    def _build(self, lato):
        outer=ttk.Frame(self.window,padding=12); outer.pack(fill='both',expand=True)
        ttk.Label(outer,text='Modifica foto — '+lato,font=('Segoe UI',16,'bold')).pack(anchor='w')
        ttk.Label(outer,text='Lavori su una copia. L’immagine originale non viene modificata.',wraplength=1000).pack(anchor='w',pady=(2,8))
        body=ttk.Frame(outer); body.pack(fill='both',expand=True)
        self.tabs=ttk.Notebook(body,width=350);self.tabs.pack(side='left',fill='y')
        self._build_crop(); self._build_color(); self._build_export()
        view=ttk.Frame(body,padding=(12,0,0,0));view.pack(side='left',fill='both',expand=True)
        top=ttk.Frame(view);top.pack(fill='x')
        ttk.Label(top,text='Anteprima nel biglietto',font=('Segoe UI',12,'bold')).pack(side='left')
        ttk.Checkbutton(top,text='Mostra senza effetti colore',variable=self.comparison).pack(side='right')
        ttk.Label(view,textvariable=self.preview_info,wraplength=690).pack(anchor='w',pady=(4,4))
        self.canvas=tk.Canvas(view,background='#e5e9ed',highlightthickness=0)
        self.canvas.pack(fill='both',expand=True)
        self.canvas.bind('<Configure>',lambda e:self.schedule())
        ttk.Label(view,textvariable=self.status,foreground='#765000',wraplength=700,justify='left').pack(anchor='w',pady=(6,0))
        actions=ttk.Frame(outer);actions.pack(fill='x',pady=(10,0))
        ttk.Button(actions,text='Ripristina regolazioni',command=self.reset).pack(side='left')
        ttk.Button(actions,text='Usa questa foto',command=self.apply).pack(side='right')
        ttk.Button(actions,text='Annulla',command=self.close).pack(side='right',padx=8)
        ttk.Button(actions,text='Esporta copia…',command=self.export).pack(side='right')

    def _scale(self, parent, row, label, key, lo, hi, suffix=''):
        ttk.Label(parent,text=label).grid(row=row,column=0,sticky='w',pady=3)
        scale=ttk.Scale(parent,from_=lo,to=hi,variable=self.variables[key],orient='horizontal')
        scale.grid(row=row,column=1,sticky='ew',padx=6)
        value=ttk.Label(parent,textvariable=self.variables[key],width=5,anchor='e')
        value.grid(row=row,column=2,sticky='e')
        if suffix:
            ttk.Label(parent,text=suffix).grid(row=row,column=3,sticky='w')

    def _build_crop(self):
        tab=ttk.Frame(self.tabs,padding=12);self.tabs.add(tab,text='Inquadratura')
        ttk.Label(tab,text='Adatta allo spazio della foto',font=('Segoe UI',10,'bold')).grid(row=0,column=0,columnspan=4,sticky='w')
        ttk.Combobox(tab,textvariable=self.variables['modalita'],state='readonly',values=(RIEMPI,CONTIENI,ALLUNGA),width=34).grid(row=1,column=0,columnspan=4,sticky='ew',pady=6)
        self._scale(tab,2,'Zoom','zoom',100,300,'%')
        self._scale(tab,3,'Posizione orizz.','posizione_x',0,100,'%')
        self._scale(tab,4,'Posizione vert.','posizione_y',0,100,'%')
        ttk.Label(tab,text='Rotazione a destra').grid(row=5,column=0,sticky='w',pady=(9,0))
        ttk.Combobox(tab,textvariable=self.variables['rotazione'],state='readonly',values=(0,90,180,270),width=6).grid(row=5,column=1,sticky='w',pady=(9,0))
        ttk.Checkbutton(tab,text='Specchia orizzontalmente',variable=self.variables['specchia']).grid(row=6,column=0,columnspan=4,sticky='w',pady=3)
        ttk.Checkbutton(tab,text='Capovolgi verticalmente',variable=self.variables['capovolgi']).grid(row=7,column=0,columnspan=4,sticky='w',pady=3)
        ttk.Button(tab,text='Colore dei bordi…',command=self.choose_background).grid(row=8,column=0,columnspan=4,sticky='w',pady=(5,0))
        ttk.Label(tab,text='I bordi si vedono scegliendo “Mostra tutta la foto”. Allungare può deformare l’immagine.',wraplength=315,justify='left').grid(row=9,column=0,columnspan=4,sticky='w',pady=8)
        tab.columnconfigure(1,weight=1)

    def _build_color(self):
        tab=ttk.Frame(self.tabs,padding=12);self.tabs.add(tab,text='Colori ed effetti')
        ttk.Label(tab,text='Effetto di partenza',font=('Segoe UI',10,'bold')).grid(row=0,column=0,columnspan=4,sticky='w')
        ttk.Combobox(tab,textvariable=self.variables['effetto'],state='readonly',values=EFFETTI).grid(row=1,column=0,columnspan=4,sticky='ew',pady=6)
        for row,args in enumerate((('Luminosità','luminosita',0,200,'%'),('Contrasto','contrasto',0,200,'%'),
                                   ('Saturazione','saturazione',0,200,'%'),('Nitidezza','nitidezza',0,200,'%'),
                                   ('Temperatura','temperatura',-100,100,''),('Tonalità','tonalita',-180,180,'°'),
                                   ('Vignettatura','vignettatura',0,100,'%')),start=2):
            self._scale(tab,row,*args)
        ttk.Button(tab,text='Ripristina solo colori',command=self.reset_colors).grid(row=9,column=0,columnspan=4,sticky='w',pady=10)
        ttk.Label(tab,text='Le regolazioni si combinano con l’effetto scelto. Usa l’anteprima per confrontarle con l’originale.',wraplength=315,justify='left').grid(row=10,column=0,columnspan=4,sticky='w')
        tab.columnconfigure(1,weight=1)

    def _build_export(self):
        tab=ttk.Frame(self.tabs,padding=12);self.tabs.add(tab,text='Dimensioni e formato')
        ttk.Label(tab,text='Dimensioni immagine',font=('Segoe UI',10,'bold')).grid(row=0,column=0,columnspan=3,sticky='w')
        ttk.Label(tab,text='Larghezza (pixel)').grid(row=1,column=0,sticky='w',pady=5)
        ttk.Entry(tab,textvariable=self.variables['larghezza'],width=9).grid(row=1,column=1,sticky='w')
        ttk.Label(tab,text='Altezza (pixel)').grid(row=2,column=0,sticky='w',pady=5)
        ttk.Entry(tab,textvariable=self.variables['altezza'],width=9).grid(row=2,column=1,sticky='w')
        ttk.Checkbutton(tab,text='Mantieni le proporzioni del biglietto',variable=self.variables['proporzioni']).grid(row=3,column=0,columnspan=3,sticky='w',pady=4)
        ttk.Button(tab,text='Imposta dimensioni consigliate',command=self.recommended).grid(row=4,column=0,columnspan=3,sticky='w',pady=4)
        ttk.Label(tab,textvariable=self.size_info,wraplength=310,justify='left').grid(row=5,column=0,columnspan=3,sticky='w',pady=(2,9))
        ttk.Separator(tab).grid(row=6,column=0,columnspan=3,sticky='ew',pady=4)
        ttk.Label(tab,text='Formato copia esportata',font=('Segoe UI',10,'bold')).grid(row=7,column=0,columnspan=3,sticky='w',pady=(5,2))
        ttk.Combobox(tab,textvariable=self.variables['formato'],state='readonly',values=tuple(FORMATI),width=12).grid(row=8,column=0,sticky='w')
        ttk.Label(tab,text='Qualità JPEG/WebP').grid(row=9,column=0,sticky='w',pady=6)
        ttk.Spinbox(tab,from_=10,to=100,textvariable=self.variables['qualita'],width=7).grid(row=9,column=1,sticky='w')
        ttk.Label(tab,text='DPI della copia').grid(row=10,column=0,sticky='w',pady=6)
        ttk.Spinbox(tab,from_=72,to=1200,increment=10,textvariable=self.variables['dpi'],width=7).grid(row=10,column=1,sticky='w')
        ttk.Label(tab,text='I DPI sono metadati per la copia esportata: ingrandire i pixel non aggiunge dettaglio alla foto.',wraplength=310,justify='left').grid(row=11,column=0,columnspan=3,sticky='w',pady=8)
        ttk.Button(tab,text='Esporta copia…',command=self.export).grid(row=12,column=0,columnspan=3,sticky='w')
        tab.columnconfigure(0,weight=1)

    def choose_background(self):
        color=colorchooser.askcolor(color=self.variables['sfondo'].get(),title='Colore dei bordi',parent=self.window)[1]
        if color:
            self.variables['sfondo'].set(color)

    def changed(self,key=None,*args):
        if self.updating:
            return
        if key in ('larghezza','altezza') and self.variables['proporzioni'].get():
            try:
                val=int(self.variables[key].get())
                if val > 0:
                    other='altezza' if key=='larghezza' else 'larghezza'
                    other_val=round(val/self.box_ratio) if key=='larghezza' else round(val*self.box_ratio)
                    self.updating=True
                    self.variables[other].set(max(1,other_val))
            except (ValueError,tk.TclError):
                pass
            finally:
                self.updating=False
        if key=='proporzioni' and self.variables['proporzioni'].get():
            self.changed('larghezza')
            return
        self.schedule()

    def schedule(self):
        if self.pending:
            self.window.after_cancel(self.pending)
        self.pending=self.window.after(180,self.refresh)

    def get_settings(self):
        try:
            settings={key:var.get() for key,var in self.variables.items()}
            for key in ('larghezza','altezza','zoom','posizione_x','posizione_y','rotazione','luminosita','contrasto',
                        'saturazione','nitidezza','temperatura','tonalita','vignettatura','qualita','dpi'):
                settings[key]=int(settings[key])
            validate_settings(settings)
            return settings
        except (ValueError,tk.TclError) as error:
            raise ValueError(str(error) if isinstance(error,ValueError) else 'Inserisci numeri validi nelle dimensioni e nelle regolazioni.') from error

    def refresh(self):
        self.pending=None
        try:
            settings=self.get_settings()
            dpi=stima_dpi(self.originale,settings,self.box_ratio,self.inches)
            self.size_info.set(f'Per {settings["dpi"]} DPI sul biglietto: circa {round(self.inches[0]*settings["dpi"])} × {round(self.inches[1]*settings["dpi"])} px.')
            hints=[f'Qualità stimata sul biglietto: {dpi:.0f} DPI effettivi.']
            if dpi<149.5: hints.append('Risoluzione bassa: la stampa potrebbe essere sgranata.')
            elif dpi<299.5: hints.append('Sotto i 300 DPI consigliati: potrebbe perdere nitidezza.')
            ratio=settings['larghezza']/settings['altezza']
            if abs(ratio/self.box_ratio-1)>.015:
                hints.append('Le proporzioni non coincidono: nel biglietto i bordi verranno ritagliati al centro.')
            if settings['modalita']==ALLUNGA:
                hints.append('Attenzione: la modalità Allunga può deformare la foto.')
            self.status.set(' '.join(hints))
            self.preview_info.set(f'{self.nome} · originale {self.originale.width} × {self.originale.height} px · uscita {settings["larghezza"]} × {settings["altezza"]} px')
            if self.comparison.get():
                settings=dict(settings,luminosita=100,contrasto=100,saturazione=100,nitidezza=100,temperatura=0,tonalita=0,vignettatura=0,effetto='Nessuno')
            canvas_w=max(80,self.canvas.winfo_width()-30);canvas_h=max(80,self.canvas.winfo_height()-30)
            image,_,_=elabora(self.preview_source,settings,(min(canvas_w,900),min(canvas_h,750)))
            # La copia esportata può avere altre proporzioni; il PDF ritaglia al centro.
            target_w=max(1,min(image.width,round(image.height*self.box_ratio)))
            target_h=max(1,min(image.height,round(image.width/self.box_ratio)))
            image=ImageOps.fit(image,(target_w,target_h),method=Image.Resampling.LANCZOS)
            self.preview_image=ImageTk.PhotoImage(image,master=self.window)
            self.canvas.delete('all')
            self.canvas.create_image((canvas_w+30)//2,(canvas_h+30)//2,image=self.preview_image,anchor='center')
        except Exception as error:
            self.preview_info.set('Controlla i valori nella scheda Dimensioni e formato.')
            self.status.set(str(error))

    def reset_colors(self):
        self.updating=True
        try:
            for key,value in (('luminosita',100),('contrasto',100),('saturazione',100),('nitidezza',100),('temperatura',0),('tonalita',0),('vignettatura',0),('effetto','Nessuno')):
                self.variables[key].set(value)
        finally:
            self.updating=False
        self.schedule()

    def reset(self):
        defaults=parametri_iniziali((round(self.inches[0]*300),round(self.inches[1]*300)))
        self.updating=True
        try:
            for key,value in defaults.items(): self.variables[key].set(value)
            self.comparison.set(False)
        finally:
            self.updating=False
        self.schedule()

    def recommended(self):
        try:
            dpi=int(self.variables['dpi'].get())
            if not 72<=dpi<=1200: raise ValueError()
            self.updating=True
            self.variables['larghezza'].set(round(self.inches[0]*dpi))
            self.variables['altezza'].set(round(self.inches[1]*dpi))
        except (ValueError,tk.TclError):
            messagebox.showerror('DPI non validi','Inserisci un valore da 72 a 1200.',parent=self.window)
        finally:
            self.updating=False
        self.schedule()

    def run_job(self,settings,callback,path=None):
        dialog=tk.Toplevel(self.window);dialog.withdraw();dialog.title('Elaborazione foto');dialog.transient(self.window)
        dialog.resizable(False,False);dialog.protocol('WM_DELETE_WINDOW',lambda:None)
        panel=ttk.Frame(dialog,padding=18);panel.pack(fill='both',expand=True)
        ttk.Label(panel,text='Elaborazione della foto in corso…',font=('Segoe UI',10,'bold')).pack(anchor='w')
        bar=ttk.Progressbar(panel,mode='indeterminate',length=300);bar.pack(fill='x',pady=8);bar.start(12)
        ttk.Label(panel,text='Attendi qualche secondo.').pack(anchor='w')
        dialog.update_idletasks();self.window.update_idletasks()
        x=max(0,self.window.winfo_rootx()+(self.window.winfo_width()-dialog.winfo_reqwidth())//2)
        y=max(0,self.window.winfo_rooty()+(self.window.winfo_height()-dialog.winfo_reqheight())//2)
        dialog.geometry(f'+{x}+{y}');dialog.deiconify();dialog.grab_set()
        messages=queue.Queue()
        def work():
            try:
                image,sx,sy=elabora(self.originale,settings)
                if path is not None: salva_immagine(path,image,settings)
                messages.put((True,(image,sx,sy)))
            except Exception as error:
                messages.put((False,str(error)))
        def check():
            try: ok,value=messages.get_nowait()
            except queue.Empty:
                if dialog.winfo_exists():dialog.after(80,check)
                return
            bar.stop();dialog.destroy()
            if not self.window.winfo_exists(): return
            self.window.grab_set()
            if ok: callback(value)
            else: messagebox.showerror('Impossibile elaborare la foto',value,parent=self.window)
        threading.Thread(target=work,daemon=True).start()
        dialog.after(80,check)

    def apply(self):
        try: settings=self.get_settings()
        except ValueError as error:
            messagebox.showerror('Controlla le impostazioni',str(error),parent=self.window);return
        def complete(value):
            image,sx,sy=value
            self.result=FotoModificata(self.originale,copy.deepcopy(settings),image,self.nome,self.percorso,sx,sy)
            self.close()
        self.run_job(settings,complete)

    def export(self):
        try: settings=self.get_settings()
        except ValueError as error:
            messagebox.showerror('Controlla le impostazioni',str(error),parent=self.window);return
        extension,_=FORMATI[settings['formato']]
        path=filedialog.asksaveasfilename(parent=self.window,title='Esporta una copia della foto',
                initialfile='Foto_modificata'+extension,defaultextension=extension,
                filetypes=[(settings['formato'],'*'+extension)])
        if not path:return
        if self.percorso and os.path.normcase(os.path.abspath(path))==os.path.normcase(os.path.abspath(self.percorso)):
            messagebox.showerror('Scegli un altro nome','Per conservare l’originale, salva la copia con un nome o in una cartella diversi.',parent=self.window);return
        def complete(value):
            messagebox.showinfo('Immagine esportata','Copia salvata:\n'+path+'\n\nPer usarla nel biglietto premi anche “Usa questa foto”.',parent=self.window)
        self.run_job(settings,complete,path)


def modifica_foto(parent, foto, originale, nome, percorso, box_ratio, inches, lato):
    if isinstance(foto,FotoModificata):
        originale=foto.originale
        nome,percorso=foto.nome,foto.percorso_originale
        settings=copy.deepcopy(foto.impostazioni)
    else:
        settings=parametri_iniziali((round(inches[0]*300),round(inches[1]*300)))
    editor=EditorFoto(parent,originale,nome,percorso,settings,box_ratio,inches,lato)
    editor.window.wait_window()
    return editor.result
