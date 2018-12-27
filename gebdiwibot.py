#!/usr/bin/python3

#encoding:utf-8

__author__="Ripx80"
__date__="25.06.2015"
__copyright__="Copyleft"
__version__="1.0"

import ezodf #must install it: pip install ezodf,xml
from itertools import tee, islice, chain
import re
import os.path
from datetime import datetime

#### BEGIN FUNCTIONS ####

def pre_next(some_iterable):
    prevs, items, nexts = tee(some_iterable, 3)
    prevs = chain([None], prevs)
    nexts = chain(islice(nexts, 1, None), [None])
    return zip(prevs, items, nexts)

def open_doc(fname):    
    if not os.path.isfile(fname):
        print('Error: filename does not exist: %s'%(fname,))
        exit(1)
    try:
        doc = ezodf.opendoc(fname)
    except(KeyError):
        print("Error: This is no a valid Format!") 
        exit(1)
    return doc

def analyse_filename(fname):
    t=fname.replace(" ","").split('-')
    stnr=t[1]
    sgnr=t[2].split('.')[0]    
    while stnr[0] == "0":
        stnr=stnr[1:]    
    return (sgnr,stnr)

def match_cnt_fn(doc,sgnr,stnr):
    #Check if Service Groupnr match    
    if not doc.sheets[0]['A36'].value:
        print("Error: Field not found A36. Can not found this field in document. Check your Format")
        exit(1)
    
    if sgnr != doc.sheets[0]['A36'].value.split(' ')[1] and sgnr != "0":
        print("Warning: Filename Group not match the Doc Groupname: %s:%s"%(sgnr,stnr))

    #Check if Territory Number match
    if int(stnr) != int(doc.sheets[0]['G3'].value):
        print("Warning: Filename Territory Number not match the Territory Number")    

def analyse_doc_cnt(doc,doc_list,sgnr,stnr):  
    doc_list[sgnr][stnr]={}
    kbcnt,rbcnt,blcnt,ridx=[0]*4
    for pre, row, nxt in pre_next(sheet.rows()):   
        for idx, c in enumerate(row):      
            if c.value != None:      
                if c.value == "Gebiet-Nr.":
                    for cc in row[idx+1:]:
                        if cc.value != None:                                        
                            if cc.value_type == "float": 
                                nr=int(cc.value)
                                if(int(stnr) != cc.value):
                                    print("Warning: Something is wrong with numbering")
                                break               
                elif c.value == "Anzahl der Klingeln":                           
                    #check if overwrite:
                    if 'cnt' in doc_list[sgnr][stnr]:
                        print('Warning: bell cnt will be overwrite')
                    else:
                        for cc in row[idx+1:]:
                            if cc.value != None and cc.value_type == "float": 
                                doc_list[sgnr][stnr]['cnt']=int(cc.value)
                                break    
                elif c.value == "Anzahl der Klingeln:": 
                    #this is the cnt of each site. need to verify
                    #print("ridx: %d | idx: %d"%(ridx,idx))                    
                    if cc.value != None and cc.value_type == "float" and row[idx+3:][0].value != None:                        
                        #print("blcnt: %d + %d"%(blcnt,int(row[idx+3:][0].value)))
                        blcnt=int(row[idx+3:][0].value)+blcnt
                elif c.value_type=="string" and c.value[0] == 'R' and c.value[1] == 'B' and row[idx-1].value == 'O':
                    rbcnt=rbcnt+1
                    
                elif c.value_type=="string" and c.value[0] == 'R' and c.value[1] == 'B' and row[idx-1].value != 'O':                
                    if 'rb' not in doc_list[sgnr][stnr]:
                        doc_list[sgnr][stnr]['rb']=[]
                    doc_list[sgnr][stnr]['rb'].append(row[idx+1].value)
                
                elif c.value_type=="string" and c.value[0] == 'K' and c.value[1] == 'B' and row[idx-1].value == 'O':
                    kbcnt=kbcnt+1
        ridx=ridx+1 
        
    if kbcnt != 0: doc_list[sgnr][stnr]['kbcnt']=kbcnt    
    if rbcnt != 0: doc_list[sgnr][stnr]['rbcnt']=rbcnt    
    if 'rb' in doc_list[sgnr][stnr]: 
        doc_list[sgnr][stnr]['rb']=list(set(doc_list[sgnr][stnr]['rb']))            
    elif rbcnt != 0:
        print("Warning: rbcnt is set but you have no rb names.")
    
    if 'cnt' not in doc_list[sgnr][stnr]:
        print("Warning: Nr. %s found but has no bell informations"%(stnr,))
    else:
        if doc_list[sgnr][stnr]['cnt'] != blcnt:
            print("Warning: Something is wrong with your bell counter: %s | %s"%(doc_list[sgnr][stnr]['cnt'],blcnt))
            #exit(1)
    doc_list[sgnr][stnr]['image']=doc.meta.count['image']
    return doc_list

def get_fname_list(fdir):
    try:
        d=os.listdir(fdir)
    except(FileNotFoundError):
        print("Error: Directory doesn't exist: %s"%(fdir,))
        exit(1)        
    if not d:
        print("Warning: Empty Directory: %s"%(fdir,))
    return sorted(d)

def refresh_overview(ovdoc,doc_list,sgnr):   
    rows=ovsheet.rows()
    idx=0
    nf=[]
    for pre, col, nxt in pre_next(ovsheet.column(0)):        
        if col.value != None:
            if col.value_type == "float":
                if str(int(col.value)) in (doc_list[sgnr]):
                    row=ovsheet.row(idx)            
                    if int(row[0].value) == int(col.value):
                        if 'cnt' in doc_list[sgnr][str(int(col.value))]:
                            row[3].set_value(int(doc_list[sgnr][str(int(col.value))]['cnt']))    #3  Bells
                        if 'kbcnt' in doc_list[sgnr][str(int(col.value))]:
                            row[14].set_value(int(doc_list[sgnr][str(int(col.value))]['kbcnt'])) #14 KB
                        if 'rbcnt' in doc_list[sgnr][str(int(col.value))]:                        
                            row[15].set_value(doc_list[sgnr][str(int(col.value))]['rbcnt']) #15 RB   
                else:
                    nf.append(int(col.value))
        idx=1+idx
    return nf

def get_doc_meta(fname,doc):
    from ezodf import meta
    print("\n### Document Meta Information###\n")
    print("FILENAME:\t\t%s"%(fname,))
    for tag in meta.TAGS.keys():
        try:
            print('%s:\t %s'%(tag.upper(),doc.meta[tag],))
        except(KeyError):
            continue
    print("\n### END Meta Information###\n")    

def set_doc_meta(doc,meta):       
    for k in meta.keys():
        doc.meta[k] = meta[k]
    doc.save()

def check_doc_pdf(fname,docdir,pdfdir,create_pdf):
    odf=(docdir+'/' if (docdir[-1]!='/') else docdir)+fname
    if os.path.isdir(pdfdir):
        pdf=(pdfdir+'/' if (pdfdir[-1]!='/') else pdfdir)+fname.replace('ods','pdf')
       
        concmd="/usr/bin/soffice --headless --convert-to pdf:calc_pdf_Export --outdir '%s' '%s'"%(pdfdir,odf)
        from subprocess import call
        if not os.path.isfile(pdf):           
            if create_pdf:
                print('Warning: no pdf file. I will create this for you: %s > %s'%(pdf,odf))
                #there are no good python lib in www for pdf converting! so i must do this :-(                
                call(concmd,shell=True) 
            else:
                print('Warning: no pdf file.')
        else:                
                odf_mime = os.stat(odf)[8]
                pdf_mime = os.stat(pdf)[8]
                if odf_mime > pdf_mime + 3600:
                    print("old pdf timestamp detected : %s / %s"%(datetime.fromtimestamp(odf_mime).strftime('%Y-%m-%d %H:%M:%S'),datetime.fromtimestamp(pdf_mime).strftime('%Y-%m-%d %H:%M:%S')))                    
                    call(concmd,shell=True)
    else:
        print('Error: This is no valid path: %s'%(pdfdir,))
        exit(1)

def genkblist(doc_list,gn):
    txt='\n\n[---- Keine Besuche im Gruppengebiet %s ----]\n\n'%(gn)   
    for key in doc_list[gn]:
        tn=doc_list[gn][key]
        if 'kbcnt' in tn:                
            txt+="Gebiet Nr: %s\tAnzahl keine Besuche: %s\t\n"%(key,tn['kbcnt'])
    return txt

def genrblist(doc_list,gn):
    txt='\n\n[---- Rückbesuche der Gruppe %s ----]\n\n'%(gn)   
    for key in doc_list[gn]:
        tn=doc_list[gn][key]
        if 'rbcnt' in tn:                
            txt+="Gebiet Nr: %s\tRückbesuche: %s\t%s\n"%(key,tn['rbcnt'],tn['rb'])
    return txt

def gentelist(doc_list,ovdoc,gn):    
    txt='\n\n[---- Länger als 6 Monate nicht bearbeitete Gebiete der Gruppe %s ----]\n\n'%(gn,)
    if gn != "0":    
        ovsheet = ovdoc.sheets[int(gn)]
    else:
        ovsheet = ovdoc.sheets[4]
    
    for idx,col in enumerate(ovsheet.column(9)):        
        if col.value != None:
            if col.value=='X':
                txt+='Gebiet Nr: %s\tzuletzt bearbeitet am: %s\n'%(int(ovsheet.column(0)[idx].value),ovsheet.column(11)[idx].value)
    txt+='\n\n[---- Länger als 1 Jahr nicht bearbeitet ----]\n\n'
    for idx,col in enumerate(ovsheet.column(10)):
        if col.value != None:
            if col.value=='X':
                txt+='Gebiet Nr: %s\tzuletzt bearbeitet am: %s\n'%(int(ovsheet.column(0)[idx].value),ovsheet.column(11)[idx].value)
    return txt

def send_mail(snd,to,subj,txt):    
    from smtplib import SMTP,SMTPException
    from email.mime.text import MIMEText
    from email.mime.multipart import MIMEMultipart

    msg = MIMEMultipart()
    msg['From'] = snd
    msg['To'] = to
    msg['Subject'] = subj
    msg.attach(MIMEText(txt,'plain','utf-8'))
    #print('Mail Header set')
    try:
        #server.set_debuglevel(1)
        server = SMTP(EMAIL_HOST,EMAIL_PORT)        
        server.starttls()
        server.ehlo()
        server.login(EMAIL_USER, EMAIL_PASSWORD)
        server.sendmail(snd, to, msg.as_string())
        server.close()
        print("Successfully sent email to: %s"%(to))
    except SMTPException:
        print("Error: unable to send email")

def get_intro():
    txt='\n== Hinweise ==\n\n'
    txt+='Hier stehen Hinweise'
    return txt

def get_banner():
    txt= '\n\n###-------------------------------###\n'
    txt+='#-----------------------------------#\n'
    txt+='#--------* Grüße von eurem *--------#\n'
    txt+='#-----------------------------------#\n'
    txt+='#--------* gebdiwi Bot :-) *--------#\n'
    txt+='###-------------------------------###\n'
    return txt
    
#### END FUNCTIONS ####

##### MAIN #####

##### BEGIN Properties ######
META={
    'title':'',
    'creator':'gebdiwi',
    'initial-creator':'gebdiwi@byom.de',
    'language':'de',
    'description':'', 
    }
    
EMAIL_HOST='smtp.gmail.com'
EMAIL_PORT=587
EMAIL_USER='gebdiwi@byom.de'
EMAIL_PASSWORD='Here is a Password req.'

# A List of Persons send to: Group:1, Group:2 aso, ov: free
EMAIL_TO={
    '0':['gebdiwi@gmail.com'],
    '1':['example@web.de','example2@web.de'],
    '2':['example@web.de','example2@web.de'],
    '3':['example@web.de','example2@web.de'],
    'ov':['gebdiwi@gmail.com']}

##### END Properties ######

##### BEGIN ARGUMENT PARSER ######
import argparse
parser = argparse.ArgumentParser()
parser.add_argument('--directory', help='set directory',required=True,nargs='+')
parser.add_argument('--overview', help='set the overview file', default=False)
parser.add_argument('--refresh', help='set the overview file', default=False,action='store_true')
parser.add_argument('--pdfdir', help='set the pdf dir. check if pdf exist with filename doc',default=False,nargs='+')
parser.add_argument('--create_pdf', help='create non exists pdfs from odf files',action='store_true',default=False)
parser.add_argument('--get_meta', help='get meta informations from odf',action='store_true',default=False)
parser.add_argument('--set_meta', help='set meta informations for odf',action='store_true',default=False)
parser.add_argument('--genrbl', help='generate a text file with a list of rbs',action='store_true',default=False)
parser.add_argument('--report', help='generate a report with rbs and territories to edit',action='store_true',default=False)
parser.add_argument('--verbose', help='print verbose infromations',action='store_true',default=False)
parser.add_argument('--mail', help='send reports by mail',action='store_true',default=False)

args = parser.parse_args()
##### END ARGUMENT PARSER ######


docdir_list = args.directory
doc_list={}

#docdir= (str(args.directory[:-1]) if args.directory[:1] == '/' else str(args.directory))

for didx,docdir in enumerate(docdir_list):
    fn_list=get_fname_list(docdir)
    for idx,docname in enumerate(fn_list):
        print("%s"%(docdir+'/'+docname))
        fnpath=docdir+'/'+docname        
        if args.pdfdir:
            if len(args.pdfdir) > didx:                
                check_doc_pdf(docname,docdir,args.pdfdir[didx],args.create_pdf)                        
        
        sgnr,stnr=analyse_filename(docname)
        if sgnr not in doc_list:
            doc_list[sgnr]={}    
        doc = open_doc(fnpath)
        doc.backup=False    
        
        if args.get_meta:
            get_doc_meta(docname,doc)
            
        if args.set_meta:
            META['title']=docname
            META['description']='Gebiet der Predigdienstgruppe %s mit der Nummer %s'%(sgnr,stnr)    
            set_doc_meta(doc,META)        
        
        match_cnt_fn(doc,sgnr,stnr)   
        sheet = doc.sheets[0]    
        analyse_doc_cnt(doc,doc_list,sgnr,stnr)


if args.refresh:
    if not args.overview:
        print("Error: You must define a overview file with --overview filename")
        exit(1)
    
    ovdoc=open_doc(args.overview)
    for gn in doc_list:
        if gn != "0":    
            ovsheet = ovdoc.sheets[int(gn)]
        else:
            ovsheet = ovdoc.sheets[4]
        nf=refresh_overview(ovdoc,doc_list,gn)  
        if nf:
            print("the following numbers not found: %s"%(nf,))
    ovdoc.save()
    print("** Update completed **")       

if args.genrbl:    
    for gn in sorted(doc_list):
        print(genrblist(doc_list,gn))


if args.report:
    if not args.overview:
        print("Error: You must define a overview file with --overview filename")
        exit(1)    
    ovdoc=open_doc(args.overview)   
    rd=datetime.now().strftime('%d-%m-%Y %H:%M')    
    txt_all=''
    for gn in sorted(doc_list):    
        txt='\n\nGebietsreport vom %s\n'%(rd,)        
        txt+=gentelist(doc_list,ovdoc,gn)
        txt+=genrblist(doc_list,gn)
        txt+=genkblist (doc_list,gn)
        txt_all+=get_intro()+txt
        if args.mail:
            txt+=get_banner()
            for mail in EMAIL_TO[gn]:
                send_mail('gebdiwi@example.com',mail,'Gebietsreport vom: %s'%(rd,),txt)            
        else:
            print(txt)
    
    if args.mail:
        txt_all+=get_banner()
        for mail in EMAIL_TO['ov']:
            send_mail('gebdiwi@example.com',mail,'Gebietsreport vom: %s'%(rd,),txt_all)

##### END MAIN #####
