import os, sys, calendar, datetime, queue, tempfile, threading, tkinter as tk
from tkinter import ttk, filedialog, messagebox
from openpyxl import load_workbook
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from PIL import Image, ImageDraw, ImageOps
from etichette import apri_etichette
from editor_foto import FotoModificata, apri_immagine, modifica_foto

BASE=os.path.dirname(os.path.abspath(__file__))
root=tk.Tk(); root.title('Biglietti di auguri - Comune'); root.geometry('1000x650')
rows=[]; current=None
source_headers=[]; source_rows=[]; source_indexes=[]
frame=ttk.Frame(root,padding=16); frame.pack(fill='both',expand=True)
ttk.Label(frame,text='Biglietti di auguri',font=('Segoe UI',20,'bold')).pack(anchor='w')
ttk.Label(frame,text='Carica un elenco Excel: il programma usa le persone presenti nel file.').pack(anchor='w',pady=(0,12))
anno_frame=ttk.Frame(frame);anno_frame.pack(fill='x',pady=(0,2))
ttk.Label(anno_frame,text='Anno del compleanno:').pack(side='left')
anno_scelto=tk.StringVar(value='')
ttk.Entry(anno_frame,textvariable=anno_scelto,width=8).pack(side='left',padx=8)
ttk.Label(anno_frame,text='Inserisci l’anno, per esempio 2026.').pack(side='left')
mesi=('Gennaio','Febbraio','Marzo','Aprile','Maggio','Giugno',
      'Luglio','Agosto','Settembre','Ottobre','Novembre','Dicembre')
mesi_attivi=[tk.BooleanVar(value=True) for _ in mesi]
month_frame=ttk.Frame(frame);month_frame.pack(fill='x',pady=(10,0))
month_header=ttk.Frame(month_frame);month_header.pack(fill='x')
ttk.Label(month_header,text='Mesi dei compleanni:').pack(side='left')
def imposta_mesi(attivi):
 for variabile in mesi_attivi: variabile.set(attivi)
 refresh()
ttk.Button(month_header,text='Tutti',command=lambda:imposta_mesi(True)).pack(side='left',padx=(12,4))
ttk.Button(month_header,text='Nessuno',command=lambda:imposta_mesi(False)).pack(side='left')
month_choices=ttk.Frame(month_frame);month_choices.pack(fill='x',pady=(4,0))
for i,(nome,variabile) in enumerate(zip(mesi,mesi_attivi)):
 ttk.Checkbutton(month_choices,text=nome,variable=variabile,command=lambda:refresh()).grid(row=i//6,column=i%6,sticky='w',padx=(0,10))
 month_choices.columnconfigure(i%6,weight=1)
year_filter=None
year_bar=ttk.Frame(frame);year_bar.pack(fill='x',pady=(9,0))
ttk.Label(year_bar,text='Anni di nascita:').pack(side='left')
year_label=tk.StringVar(value='Tutti gli anni')
ttk.Button(year_bar,text='Scegli anni…',command=lambda:scegli_anni()).pack(side='left',padx=(10,8))
ttk.Label(year_bar,textvariable=year_label).pack(side='left')
searchbar=ttk.Frame(frame); searchbar.pack(fill='x',pady=(10,0))
search=tk.StringVar(value='')
ttk.Label(searchbar,text='Cerca persona:').pack(side='left')
ttk.Entry(searchbar,textvariable=search,width=36).pack(side='left',padx=8)
ttk.Button(searchbar,text='Cancella ricerca',command=lambda:search.set('')).pack(side='left')
status=tk.StringVar(value='Nessun file caricato.')
table_frame=ttk.Frame(frame); table_frame.pack(fill='both',expand=True,pady=12)
tree=ttk.Treeview(table_frame,columns=('nome','sesso','nascita','compleanno','eta','modifica','rimuovi'),show='headings')
for c,t,w in [('nome','Cognome e nome',360),('sesso','Sesso',70),('nascita','Data nascita',120),('compleanno','Compleanno',130),('eta','Età',55),('modifica','Modifica',65),('rimuovi','Rimuovi',65)]: tree.heading(c,text=t); tree.column(c,width=w)
scrollbar=ttk.Scrollbar(table_frame,orient='vertical',command=tree.yview)
tree.configure(yscrollcommand=scrollbar.set)
scrollbar.pack(side='right',fill='y')
tree.pack(side='left',fill='both',expand=True)
ttk.Label(frame,textvariable=status,wraplength=930).pack(anchor='w')
def load():
 global rows,current,year_filter,source_headers,source_rows,source_indexes
 p=filedialog.askopenfilename(filetypes=[('Excel','*.xlsx *.xlsm')]);
 if not p:return
 try:
  workbook=load_workbook(p,read_only=True,data_only=True)
  try:
   values=iter(workbook.active.values); headers=None; header_line=0
   for header_line,row in enumerate(values,1):
    h=[str(x or '').strip().lower() for x in row]
    if 'cognome e nome' in h and 'data nascita' in h:
     headers=[str(x).strip() if x is not None and str(x).strip() else f'Campo {i+1}' for i,x in enumerate(row)]
     break
   if headers is None: raise ValueError('Non trovo le colonne “cognome e nome” e “data nascita”.')
   indexes=[h.index(x) for x in ('cognome e nome','data nascita','sesso')]
   loaded_rows=[]; loaded_source=[]; invalid=[]
   for n,row in enumerate(values,start=header_line+1):
    name,dt,sex=[row[i] if i<len(row) else None for i in indexes]
    if not name and not dt: continue
    if isinstance(dt,datetime.datetime): dt=dt.date()
    if not name or not isinstance(dt,datetime.date) or str(sex).strip().upper() not in ('M','F'):
     invalid.append(n); continue
    person=(str(name).strip(),dt,str(sex).strip().upper())
    record=list(row[:len(headers)])+[None]*max(0,len(headers)-len(row))
    for col,value in zip(indexes,person): record[col]=value
    loaded_rows.append(person);loaded_source.append(record)
  finally:
   workbook.close()
  rows=loaded_rows;source_rows=loaded_source;source_headers=headers;source_indexes=indexes
  current=os.path.basename(p); year_filter=None; search.set(''); refresh()
  if invalid: messagebox.showwarning('Righe da controllare',f'{len(invalid)} righe escluse per nome, data o sesso mancanti/non validi. Righe Excel: '+', '.join(map(str,invalid[:40])))
 except Exception as e: messagebox.showerror('Errore apertura',str(e))
def anni_disponibili():
 return sorted({dt.year for _,dt,_ in rows},reverse=True)

def anno_compleanno():
 valore=anno_scelto.get().strip()
 if len(valore)!=4 or not valore.isascii() or not valore.isdecimal(): return None
 anno=int(valore)
 return anno if 1000<=anno<=9999 else None

def corrisponde_filtri(dt):
 return mesi_attivi[dt.month-1].get() and (year_filter is None or dt.year in year_filter)

def aggiorna_etichetta_anni():
 anni=anni_disponibili()
 scelti=anni if year_filter is None else [anno for anno in anni if anno in year_filter]
 if year_filter is None or (anni and len(scelti)==len(anni)):
  year_label.set('Tutti gli anni')
 elif not scelti:
  year_label.set('Nessun anno selezionato')
 elif len(scelti)==1:
  anno=scelti[0];target=anno_compleanno()
  year_label.set(f'{anno} ({target-anno} anni)' if target is not None and target>=anno else str(anno))
 else:
  year_label.set(f'{len(scelti)} anni selezionati')

def scegli_anni():
 anni=anni_disponibili()
 if not anni:
  messagebox.showinfo('Anni di nascita','Carica prima un elenco Excel con persone valide.',parent=root)
  return
 dialog=tk.Toplevel(root);dialog.withdraw();dialog.title('Seleziona gli anni di nascita');dialog.transient(root)
 box=ttk.Frame(dialog,padding=16);box.pack(fill='both',expand=True)
 target=anno_compleanno()
 ttk.Label(box,text='Seleziona uno o più anni di nascita',font=('Segoe UI',12,'bold')).pack(anchor='w')
 spiegazione=(f'Fra parentesi trovi l’età compiuta nel {target}.' if target is not None
              else 'Inserisci l’anno del compleanno per visualizzare le età.')
 ttk.Label(box,text=spiegazione).pack(anchor='w',pady=(4,8))
 ttk.Label(box,text='Fai clic sugli anni che vuoi includere.').pack(anchor='w',pady=(0,6))
 lista_frame=ttk.Frame(box);lista_frame.pack(fill='both',expand=True)
 lista=tk.Listbox(lista_frame,selectmode='multiple',exportselection=False,height=min(12,len(anni)),width=30)
 barra=ttk.Scrollbar(lista_frame,orient='vertical',command=lista.yview)
 lista.configure(yscrollcommand=barra.set);barra.pack(side='right',fill='y');lista.pack(side='left',fill='both',expand=True)
 for i,anno in enumerate(anni):
  testo=f'{anno} ({target-anno} anni)' if target is not None and target>=anno else str(anno)
  lista.insert('end',testo)
  if year_filter is None or anno in year_filter:lista.selection_set(i)
 actions=ttk.Frame(box);actions.pack(fill='x',pady=(10,0))
 ttk.Button(actions,text='Tutti',command=lambda:lista.selection_set(0,'end')).pack(side='left')
 ttk.Button(actions,text='Nessuno',command=lambda:lista.selection_clear(0,'end')).pack(side='left',padx=6)
 def applica():
  global year_filter
  scelti={anni[int(i)] for i in lista.curselection()}
  year_filter=None if len(scelti)==len(anni) else scelti
  dialog.destroy();refresh()
 ttk.Button(actions,text='Applica',command=applica).pack(side='right')
 ttk.Button(actions,text='Annulla',command=dialog.destroy).pack(side='right',padx=6)
 root.update_idletasks();dialog.update_idletasks()
 x=max(0,root.winfo_rootx()+(root.winfo_width()-dialog.winfo_reqwidth())//2)
 y=max(0,root.winfo_rooty()+(root.winfo_height()-dialog.winfo_reqheight())//2)
 dialog.geometry(f'+{x}+{y}');dialog.deiconify();dialog.grab_set();dialog.focus_set()

def selection():
 y=anno_compleanno()
 if y is None: raise ValueError('Inserisci un anno del compleanno valido di quattro cifre.')
 result=[]
 for name,dt,sex in rows:
  if not corrisponde_filtri(dt): continue
  day=min(dt.day,calendar.monthrange(y,dt.month)[1]); birthday=datetime.date(y,dt.month,day)
  result.append((name,dt,sex,birthday,y-dt.year))
 return sorted(result,key=lambda x:(-x[1].toordinal(),x[0]))
def refresh(*args):
 tree.delete(*tree.get_children())
 try:
  aggiorna_etichetta_anni()
  y=anno_compleanno(); totale_filtrato=0; visible=0; terms=search.get().casefold().split()
  for idx in sorted(range(len(rows)),key=lambda i:(-rows[i][1].toordinal(),rows[i][0])):
   name,dt,sex=rows[idx]
   if not corrisponde_filtri(dt): continue
   totale_filtrato+=1
   if not all(term in name.casefold() for term in terms): continue
   visible+=1
   compleanno='—';eta='—'
   if y is not None and y>=dt.year:
    b=datetime.date(y,dt.month,min(dt.day,calendar.monthrange(y,dt.month)[1]))
    compleanno=b.strftime('%d/%m/%Y');eta=y-dt.year
   tree.insert('', 'end', iid=str(idx), values=(name,sex,dt.strftime('%d/%m/%Y'),compleanno,eta,'✎','✕'))
  anno_info=f'Anno: {y}.' if y is not None else 'Inserisci un anno del compleanno valido di quattro cifre.'
  status.set(f'{current or "Nessun file"} — {visible} visibili, {totale_filtrato} nei mesi e anni scelti su {len(rows)} persone. PDF: {totale_filtrato} biglietti, anche per chi è nascosto dalla ricerca. {anno_info}')
 except (ValueError,OverflowError): status.set('Impossibile calcolare la data del compleanno.')
search.trace_add('write',refresh)
anno_scelto.trace_add('write',refresh)
def azione_riga(event):
 if tree.identify_region(event.x,event.y) != 'cell': return
 item=tree.identify_row(event.y); col=tree.identify_column(event.x)
 if not item or col not in ('#6','#7'): return
 idx=int(item)
 if col=='#7':
  rows.pop(idx); source_rows.pop(idx); refresh(); return
 name,dt,sex=rows[idx]
 dialog=tk.Toplevel(root); dialog.withdraw(); dialog.title('Modifica persona'); dialog.transient(root)
 frm=ttk.Frame(dialog,padding=18); frm.pack(fill='both',expand=True)
 fields=[tk.StringVar(value=name),tk.StringVar(value=sex),tk.StringVar(value=dt.strftime('%d/%m/%Y'))]
 for label,var in zip(('Cognome e nome','Sesso (M o F)','Data di nascita (gg/mm/aaaa)'),fields):
  ttk.Label(frm,text=label).pack(anchor='w'); ttk.Entry(frm,textvariable=var,width=42).pack(fill='x',pady=(0,8))
 def save():
  try:
   newname=fields[0].get().strip(); newsex=fields[1].get().strip().upper(); newdate=datetime.datetime.strptime(fields[2].get().strip(),'%d/%m/%Y').date()
   if not newname or newsex not in ('M','F'): raise ValueError('Inserisci nome e cognome e sesso M oppure F.')
   if year_filter is not None: year_filter.add(newdate.year)
   rows[idx]=(newname,newdate,newsex)
   for col,value in zip(source_indexes,(newname,newdate,newsex)): source_rows[idx][col]=value
   dialog.destroy(); refresh()
  except ValueError as e: messagebox.showerror('Controlla i dati',str(e),parent=dialog)
 ttk.Button(frm,text='Salva modifiche',command=save).pack(anchor='e',pady=(8,0))
 root.update_idletasks(); dialog.update_idletasks()
 x=root.winfo_rootx()+(root.winfo_width()-dialog.winfo_reqwidth())//2
 y=root.winfo_rooty()+(root.winfo_height()-dialog.winfo_reqheight())//2
 dialog.geometry(f'+{max(0,x)}+{max(0,y)}')
 dialog.deiconify(); dialog.grab_set(); dialog.focus_set()
tree.bind('<ButtonRelease-1>',azione_riga)

PHOTO_BOXES=[('Fronte - foto sotto lo stemma',(710,1210,1195,1607),False),('Retro - foto sotto gli auguri',(745,201,1169,583),True)]

def leggi_foto(photo,index):
 if isinstance(photo,FotoModificata): return photo.immagine.copy()
 if photo is not None: return apri_immagine(photo)
 _,(x1,y1,x2,y2),rot=PHOTO_BOXES[index]
 with Image.open(os.path.join(BASE,'modello_scansione.png')) as source:
  w,h=source.size
  img=source.convert('RGB').crop((round(x1*w/1274),round(y1*h/1800),round(x2*w/1274),round(y2*h/1800)))
 if rot: img=img.rotate(180)
 return img

def prepara_foto(path,index):
 # Il PDF usa lo stesso adattamento mostrato in anteprima.
 _,(x1,y1,x2,y2),rot=PHOTO_BOXES[index]
 target_w,target_h=x2-x1,y2-y1
 img=leggi_foto(path,index)
 width,height=img.size; ratio=target_w/target_h
 crop_w=min(width,height*ratio); crop_h=min(height,width/ratio)
 # Dimensione effettiva della fotografia sul foglio A4 (modello PDF 595.2 × 841.44 punti).
 inches_w=(target_w/1274*595.2)/72; inches_h=(target_h/1800*841.44)/72
 sx=min(1,path.scala_x) if isinstance(path,FotoModificata) else 1
 sy=min(1,path.scala_y) if isinstance(path,FotoModificata) else 1
 effective_dpi=min(crop_w*sx/inches_w,crop_h*sy/inches_h)
 crop_pct=(1-(crop_w*crop_h)/(width*height))*100
 recommended=(round(inches_w*300),round(inches_h*300))
 preview=ImageOps.fit(img,(target_w,target_h),method=Image.Resampling.LANCZOS)
 return preview,(width,height),effective_dpi,crop_pct,recommended

def scegli_foto():
 from PIL import ImageTk
 dialog=tk.Toplevel(root); dialog.withdraw(); dialog.title('Immagini dei biglietti'); dialog.transient(root)
 dialog.geometry('900x680'); dialog.minsize(760,600)
 box=ttk.Frame(dialog,padding=16); box.pack(fill='both',expand=True)
 ttk.Label(box,text='Scegli le fotografie',font=('Segoe UI',17,'bold')).pack(anchor='w')
 ttk.Label(box,text='Le foto predefinite sono già inserite. Puoi sostituirle oppure continuare.',wraplength=850).pack(anchor='w',pady=(3,4))
 ttk.Label(box,text='Formati accettati: JPG, PNG, BMP e WebP. In Modifica foto puoi scegliere inquadratura, bordi, colori e altre regolazioni.',wraplength=850).pack(anchor='w',pady=(0,10))
 selected=[None,None]; previews=[None,None]; labels=[]; details=[]; panels=[]; accepted=[False]; warnings=[False,False]
 body=ttk.Frame(box); body.pack(fill='both',expand=True)
 for i,(title,coords,rot) in enumerate(PHOTO_BOXES):
  panel=ttk.LabelFrame(body,text=title,padding=10);panel.grid(row=0,column=i,sticky='nsew',padx=5);body.columnconfigure(i,weight=1);body.rowconfigure(0,weight=1)
  photo_label=ttk.Label(panel,anchor='center');photo_label.pack(pady=(0,6));labels.append(photo_label)
  info=tk.StringVar();details.append(info)
  ttk.Label(panel,textvariable=info,wraplength=390,justify='left').pack(anchor='w',fill='x',pady=3)
  buttons=ttk.Frame(panel);buttons.pack(anchor='w',pady=8);panels.append((panel,buttons))
 def show(i,path,alert=False):
  try:
   preview,dimensions,dpi,crop,recommended=prepara_foto(path,i)
   preview.thumbnail((360,220),Image.Resampling.LANCZOS)
   previews[i]=ImageTk.PhotoImage(preview);labels[i].configure(image=previews[i])
   w,h=dimensions;rw,rh=recommended
   warning=[]
   if path and dpi<149.5: warning.append('Risoluzione bassa: potrebbe risultare sgranata in stampa.')
   elif path and dpi<299.5: warning.append('Risoluzione inferiore a quella consigliata: potrebbe perdere nitidezza.')
   if path and crop>=5: warning.append(f'Proporzioni diverse: verrà ritagliato circa il {crop:.0f}% della foto, soprattutto ai bordi.')
   warnings[i]=bool(warning)
   selected[i]=path
   if isinstance(path,FotoModificata): text='Foto modificata: '+path.nome
   elif path: text='File: '+os.path.basename(path)
   else: text='Foto predefinita del modello'
   text+=f'\nDimensioni: {w} × {h} px. Consigliati circa {rw} × {rh} px (300 dpi).'
   text+=f'\nLa foto viene centrata e ritagliata per riempire lo spazio. Anteprima sopra.'
   text+='\n'+'\n'.join(warning) if warning else ('\nNessun problema di qualità rilevato.' if path else '\nImmagine predefinita del nuovo modello.')
   details[i].set(text)
   if warning and alert: messagebox.showwarning('Controlla la fotografia','\n'.join(warning)+'\n\nPuoi comunque continuare e usare questa foto.',parent=dialog)
  except Exception as e:
   messagebox.showerror('Immagine non valida',f'Non riesco ad aprire questa immagine: {e}',parent=dialog)
 def choose(i):
  path=filedialog.askopenfilename(parent=dialog,title='Scegli '+PHOTO_BOXES[i][0],filetypes=[('Immagini','*.jpg *.jpeg *.png *.bmp *.webp')])
  if path: show(i,path,True)
 def edit(i):
  try:
   current=selected[i]
   original=current.originale if isinstance(current,FotoModificata) else leggi_foto(current,i)
   name=current.nome if isinstance(current,FotoModificata) else os.path.basename(current) if current else 'Foto predefinita del modello'
   source_path=current.percorso_originale if isinstance(current,FotoModificata) else (current or os.path.join(BASE,'modello_scansione.png'))
   title,(x1,y1,x2,y2),_=PHOTO_BOXES[i]
   box_w,box_h=x2-x1,y2-y1
   inches=(box_w/1274*595.2/72,box_h/1800*841.44/72)
   modified=modifica_foto(dialog,current,original,name,source_path,box_w/box_h,inches,'Fronte' if i==0 else 'Retro')
   if modified is not None: show(i,modified,True)
  except Exception as e:
   messagebox.showerror('Impossibile modificare la foto',str(e),parent=dialog)
 for i,(panel,buttons) in enumerate(panels):
  ttk.Button(buttons,text='Scegli foto',command=lambda i=i:choose(i)).pack(side='left',padx=(0,6))
  ttk.Button(buttons,text='Modifica foto',command=lambda i=i:edit(i)).pack(side='left',padx=(0,6))
  ttk.Button(buttons,text='Ripristina predefinita',command=lambda i=i:show(i,None)).pack(side='left')
  show(i,None)
 actions=ttk.Frame(box);actions.pack(fill='x',pady=(10,0))
 def done():
  if any(warnings) and not messagebox.askyesno('Usare comunque le foto?', 'Una o più immagini hanno avvisi di qualità o ritaglio. Vuoi comunque creare il PDF con queste foto?',parent=dialog): return
  accepted[0]=True;dialog.destroy()
 ttk.Button(actions,text='Indietro',command=dialog.destroy).pack(side='left')
 ttk.Button(actions,text='Crea PDF',command=done).pack(side='right')
 root.update_idletasks();dialog.update_idletasks()
 x=max(0,root.winfo_rootx()+(root.winfo_width()-900)//2);y=max(0,root.winfo_rooty()+(root.winfo_height()-680)//2)
 dialog.geometry(f'900x680+{x}+{y}');dialog.deiconify();dialog.grab_set();dialog.focus_set()
 dialog.wait_window()
 return selected if accepted[0] else None


def crea_pdf(path, data, foto, progresso=None):
 # Il modello è ricavato dal nuovo biglietto, già ripulito dai dati di esempio.
 with Image.open(os.path.join(BASE,'modello_scansione.png')) as source:
  img=source.convert('RGB')
 w,h=img.size
 for index,photo in enumerate(foto):
  if photo is None: continue
  _,(x1,y1,x2,y2),rot=PHOTO_BOXES[index]
  x1,x2=round(x1*w/1274),round(x2*w/1274)
  y1,y2=round(y1*h/1800),round(y2*h/1800)
  source=leggi_foto(photo,index)
  fitted=ImageOps.fit(source,(x2-x1,y2-y1),method=Image.Resampling.LANCZOS)
  if rot: fitted=fitted.rotate(180)
  img.paste(fitted,(x1,y1))
 bg=ImageReader(img)
 W,H=595.32,841.92
 font='Times-BoldItalic'
 c=canvas.Canvas(path,pagesize=(W,H))
 c.beginForm('modello',0,0,W,H)
 c.drawImage(bg,0,0,width=W,height=H)
 c.endForm()
 blue=(0,0.40,1)
 def lines_for(text, width, size):
  lines=[]; line=''
  for word in text.split():
   candidate=(line+' '+word).strip()
   if line and pdfmetrics.stringWidth(candidate,font,size)>width:
    lines.append(line); line=word
   else: line=candidate
  if line: lines.append(line)
  return lines
 def copertina(name,sex):
  prefix='Gent.le sig.ra ' if sex=='F' else 'Gent.le sig. '
  left,right=348,542
  size=13
  while size>=7:
   indent=left+pdfmetrics.stringWidth(prefix,font,size)
   lines=lines_for(name,right-indent,size)
   if len(lines)<=3 and all(pdfmetrics.stringWidth(line,font,size)<=right-indent for line in lines): break
   size-=0.5
  else: raise ValueError('Nominativo troppo lungo per il riquadro: '+name)
  c.setFillColorRGB(*blue); c.setFont(font,size)
  c.drawString(left,H-785,prefix)
  for i,line in enumerate(lines): c.drawString(indent,H-785-i*(size+2),line)
 def data_ruotata(text,x):
  c.saveState();c.setFillColorRGB(*blue);c.setFont('Times-Italic',13)
  c.translate(x,H-387);c.rotate(180);c.drawCentredString(0,0,text);c.restoreState()
 for index,(name,dt,sex,b,age) in enumerate(data,1):
  c.doForm('modello')
  copertina(name,sex)
  data_ruotata(dt.strftime('%d.%m.%Y'),209)
  data_ruotata(b.strftime('%d.%m.%Y'),81)
  c.showPage()
  if progresso: progresso(index)
 if progresso: progresso(None)
 c.save()

def mostra_avanzamento(path,data,foto):
 dialog=tk.Toplevel(root)
 dialog.withdraw(); dialog.title('Creazione PDF'); dialog.transient(root)
 dialog.resizable(False,False)
 dialog.protocol('WM_DELETE_WINDOW',lambda: None)
 box=ttk.Frame(dialog,padding=20); box.pack(fill='both',expand=True)
 ttk.Label(box,text='Creazione dei biglietti in corso',font=('Segoe UI',12,'bold')).pack(anchor='w')
 text=tk.StringVar(value='Preparazione del modello…')
 ttk.Label(box,textvariable=text).pack(anchor='w',pady=(8,8))
 bar=ttk.Progressbar(box,orient='horizontal',mode='determinate',maximum=len(data)+1,length=370)
 bar.pack(fill='x')
 ttk.Label(box,text='Attendi il messaggio di completamento.').pack(anchor='w',pady=(8,0))
 dialog.update_idletasks();root.update_idletasks()
 x=max(0,root.winfo_rootx()+(root.winfo_width()-dialog.winfo_reqwidth())//2)
 y=max(0,root.winfo_rooty()+(root.winfo_height()-dialog.winfo_reqheight())//2)
 dialog.geometry(f'+{x}+{y}');dialog.deiconify();dialog.grab_set()
 events=queue.Queue()
 def work():
  temporary=None
  try:
   descriptor,temporary=tempfile.mkstemp(prefix='.biglietti_',suffix='.pdf',dir=os.path.dirname(os.path.abspath(path)))
   os.close(descriptor)
   crea_pdf(temporary,data,foto,lambda count: events.put(('progress',count)))
   os.replace(temporary,path)
   temporary=None
   events.put(('done',None))
  except Exception as error:
   events.put(('error',str(error)))
  finally:
   if temporary is not None:
    try: os.unlink(temporary)
    except OSError: pass
 def controlla():
  try:
   while True:
    kind,value=events.get_nowait()
    if kind=='progress':
     if value is None:
      text.set('Salvataggio del PDF in corso…')
     else:
      bar['value']=value
      text.set(f'Biglietti preparati: {value} su {len(data)}')
    elif kind=='done':
     bar['value']=len(data)+1
     dialog.destroy()
     messagebox.showinfo('PDF creato',f'Creati {len(data)} biglietti.\n\nPrima di stampare tutto, verifica una pagina: stampa A4 al 100% e controlla piegatura e margini.',parent=root)
     return
    else:
     dialog.destroy()
     messagebox.showerror('Impossibile creare il PDF',value,parent=root)
     return
  except queue.Empty:
   dialog.after(75,controlla)
 def avvia():
  threading.Thread(target=work,daemon=True).start()
  controlla()
 dialog.after(80,avvia)


def generate():
 try:
  if anno_compleanno() is None: raise ValueError('Inserisci un anno del compleanno valido di quattro cifre.')
  data=selection()
  if not any(mese.get() for mese in mesi_attivi): raise ValueError('Seleziona almeno un mese per creare i biglietti.')
  if year_filter is not None and not year_filter: raise ValueError('Seleziona almeno un anno di nascita per creare i biglietti.')
  if not data: raise ValueError('Nessuna persona da stampare con i mesi e gli anni selezionati. Controlla i filtri o carica un elenco Excel.')
  if any(b.year<d.year for _,d,_,b,_ in data):
   raise ValueError('L’anno del compleanno è precedente alla nascita di almeno una persona selezionata. Controlla l’anno o i filtri.')
  if any(d.month==2 and d.day==29 and b.day==28 for _,d,_,b,_ in data):
   if not messagebox.askyesno('Compleanni del 29 febbraio','In un anno non bisestile i compleanni del 29 febbraio saranno indicati al 28 febbraio. Continuare?'): return
  foto=scegli_foto()
  if foto is None: return
  path=filedialog.asksaveasfilename(defaultextension='.pdf',initialfile=('Biglietti_'+os.path.splitext(current)[0]+'.pdf' if current else 'Biglietti.pdf'),filetypes=[('PDF','*.pdf')])
  if not path:return
  mostra_avanzamento(path,data,foto)
 except Exception as e: messagebox.showerror('Impossibile creare il PDF',str(e))
def apri_finestra_etichette():
 if not rows:
  messagebox.showinfo('Crea etichette','Carica prima un elenco Excel con persone valide.',parent=root)
  return
 indici=sorted((i for i,(_,dt,_) in enumerate(rows) if corrisponde_filtri(dt)),
               key=lambda i:(-rows[i][1].toordinal(),rows[i][0]))
 if not indici:
  messagebox.showinfo('Crea etichette','Nessuna persona corrisponde ai mesi e anni di nascita selezionati.',parent=root)
  return
 dati=[list(source_rows[i]) for i in indici]
 apri_etichette(root,source_headers,dati,current)

buttons=ttk.Frame(frame); buttons.pack(fill='x',pady=10)
ttk.Button(buttons,text='Carica Excel',command=load).pack(side='left')
ttk.Button(buttons,text='Aggiorna elenco',command=refresh).pack(side='left',padx=8)
ttk.Button(buttons,text='Avanti',command=generate).pack(side='right')
ttk.Button(buttons,text='Crea etichette / PDF',command=apri_finestra_etichette).pack(side='right',padx=(0,8))
root.mainloop()
