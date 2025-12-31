import array
import sys
import os
import re
import json
import random
import csv
import pytesseract
import fitz  # PyMuPDF
import sys
import os
from pathlib import Path
from datetime import datetime
import xml.etree.ElementTree as ET
from PIL import Image, ImageDraw
from pypdf import PdfReader, PdfWriter
from pdf_extract_kit.registry.registry import TASK_REGISTRY
from pdf_extract_kit.utils.data_preprocess import load_pdf
from pdf_extract_kit.tasks.base_task import BaseTask
from pdf_extract_kit.utils.pdf_utils import save_pdf
from pdf_extract_kit.tasks.ocr.emails import Email
from subprocess import run

@TASK_REGISTRY.register("ocr")
class OCRTask(BaseTask):
    score_valid = 0.75 # puntuació mínima per considerar un text
    marge_posicio_x = 15 #+ 
    marge_posicio_y = 12 #+
    marge_document_y = 0
    marge_document_x = 0
    dadesGrupNegoci = []
    logfilename = 'outputs/log.txt'
    reglog=''
    llistplanes=[]
    cuadresplantilla=[]
    jsondades=[]
    numPlanes=0
    documents=0
    correctes=0
    revisions=0
    eliminats=0
    nomDocumentProcessat=''
    
    def __init__(self, model):
        self.carregaDadesNavision()
        """init the task based on the given model.
        
        Args:
            model: task model, must contains predict function.
        """
        super().__init__(model)

    def predict_image(self, image):
        """predict on one image, reture text detection and recognition results.
        
        Args:
            image: PIL.Image.Image, (if the model.predict function support other types, remenber add change-format-function in model.predict)
            
        Returns:
            List[dict]: list of text bbox with it's content
            
        Return example:
            [
                {
                    "category_type": "text",
                    "poly": [
                        380.6792698635707,
                        159.85058512958923,
                        765.1419999999998,
                        159.85058512958923,
                        765.1419999999998,
                        192.51073013642917,
                        380.6792698635707,
                        192.51073013642917
                    ],
                    "text": "this is an example text",
                    "score": 0.97
                },
                ...
            ]
        """
        return self.model.predict(image)

        
        
    def prepare_input_files(self, input_path):
        if os.path.isdir(input_path):
            file_list = [os.path.join(Path(input_path), fname) for fname in os.listdir(input_path)]
        else:
            file_list = [input_path]
        return file_list
            
    def process(self, input_path, save_dir=None, visualize=False):
        #opcions de process
        sw_drivebaixada=0
        sw_drivepujada=0
        sw_separarPlanes=1
        sw_nomesSeparar=0       
        sw_guardarjson=1
        sw_guardarjsonerronis=1
        ult_camps=[]
        if(os.path.isdir('/srv/pdf-extract')):
            directoriDrive ="/srv/pdf-extract/PDF-Extract-Kit-main/assets/inputs/ocr"
            separador='/'
        else:
            if (os.path.isdir(r'c:\PDF-Extract-Kit-main\PDF-Extract-Kit-main')):  
                directoriDrive =r"c:\PDF-Extract-Kit-main\PDF-Extract-Kit-main\assets\inputs\ocr"                     
                separador='\\'
            else:      
                directoriDrive ="/mnt/c/PDF-Extract-Kit-main/PDF-Extract-Kit-main/assets/inputs/ocr"
                separador='/'
       
        if(sw_drivebaixada):
            run(["rclone", "move", "ocr_input:", directoriDrive])
        dir_list= self.prepare_input_files(input_path)
        # print(dir_list)
        self.guardar_logs('Comença el procés ',0,0,'')
        for directori in dir_list:
            dirpos=directori.rfind(separador)+1
            plataforma=directori[dirpos:100]
            print('Plataforma: ',plataforma)                    
            ##### Asignar arxiu LOG################
            dataprocess= datetime.now().strftime("%Y-%m-%d")
            self.logfilename='outputs/log_'+plataforma+'_'+dataprocess+'.txt'
            
            ##### Agafa els PDF's de cada plataforma
            file_list = self.prepare_input_files(directori)
            
            res_list = []
            # if res_list != []:
            #     print('--------------------------------------------------')
            #     print(file_list)
            #     print('--------------------------------------------------')  
            carpeta=save_dir

            if(sw_separarPlanes):    
                # print('Separa planes '+datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
                for prefpath in file_list:  # Per cada PDF
                    if prefpath.endswith(".pdf") or prefpath.endswith(".PDF"):
                        tot_planes=self.separa_i_orienta(prefpath)  
                        if(tot_planes>1):
                            self.guardar_logs('Separar planes ('+directori+')',0,1,prefpath,tot_planes)
                if(sw_nomesSeparar):
                      sys.exit()  
                file_list = self.prepare_input_files(directori)  
                
            self.numPlanes=0     
            self.documents=0
            self.correctes=0
            self.revisions=0
            self.eliminats=0
            for fpath in file_list:  # Per cada PDF
                if fpath.endswith(".pdf") or fpath.endswith(".PDF"):
                    basename = os.path.basename(fpath)[:-4].replace('.','_') # treiem els punts intermitjos del nom
                    self.nomDocumentProcessat=basename
                    images = load_pdf(fpath)  #carrega les imatges del PDF
                    nom=''
                    pdf_res = []
                    #print('Comença document '+basename+' '+datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
                    
                    print('--------------------------------------------------')
                    print(basename)
                    print('--------------------------------------------------')  
                    for page, img in enumerate(images):  #per cada imatge
                        # self.reglog= datetime.now().strftime("%Y-%m-%d %H:%M:%S")+' '  
                        # self.reglog+=basename+' '+str(page+1)+ '\r\n\t'
                        self.numPlanes=self.numPlanes+1
                        #self.guardar_logs('Processant '+basename,0,1,basename + f"_{page+1}.pdf")
                        self.reglog = '' 
                        print ('entrar a predict_image')
                        page_res = self.predict_image(img)
                        pdf_res.append(page_res)
                        print ('pause')
                        nom = save_dir + separador + basename
                        ClientNav=None
                        #print(save_dir,'------------------------->>> ')
                        sw_error=False   #Guardará la imatge si hi ha error en cas d'error i sigui =1

                        if save_dir:
                            camps=[]
                            plant=''
                            dadesJsonRevisat=[]
                            fitxerJson=directori+separador+basename+'.json'
                            
                            if os.path.exists(fitxerJson):
                                try:
                                    with open(fitxerJson, 'r', encoding='utf-8') as file:
                                        dadesJsonRevisat= json.load(file)
                                    if (dadesJsonRevisat !=''):
                                        plant= dadesJsonRevisat['plantilla']
                                        camps.append(dadesJsonRevisat['codiclient'])
                                        camps.append(dadesJsonRevisat['nomclient'])
                                        camps.append(dadesJsonRevisat['dataalbara'])
                                        camps.append(dadesJsonRevisat['numalbara'])
                                        #print('Dades del JSON : ',plant,camps)
                                        self.reglog = 'Dades del JSON : '+plant+str(camps)
                                except:
                                    pass
                                os.remove(fitxerJson)
                            #print('Plantilla detectada : ',plant)
                            # activar la linia en cas de voler guardar el json per veure les posicions del camps
                            if(sw_guardarjson):
                                self.save_json_result( page_res, os.path.join(save_dir,'jsons', basename + f"_{page+1}.json"))
                            if(plant==''):  
                                #print('Detectant plantilla '+basename+' '+datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
                                plantilles=self.detectarPlantilla(page_res)
                                if (plantilles==[]):
                                    sw_error=True
                                else:
                                    plant=plantilles[0]
                                    arxplantilla=plantilles[1]

                            #print('--------------------------------------------------')
                            #print('Plana : ',page+1)
                            print ('La plantilla es: ',plant)

                            ####### eliminar documents en cas de plantilla XX
                            if (plant=='XX'):
                                self.guardar_logs('Document eliminat XX',0,1, basename + f"_{page+1}.pdf")
                                self.eliminats=self.eliminats+1
                                continue

                            if(plant is not None and sw_error==False):
                                if (camps)==[]:
                                    #print('Detectant camps '+basename+' '+datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
                                    camps=self.detectarCamps(arxplantilla,page_res)
                                #print ('Els camps son: ',camps)  
                                x=0  
                                x=len(camps)
                                #print ('Resultats de camps : ',x)
                                if (plant=='VP'):
                                    if(ult_camps!=[]):
                                        if (x<4 and camps[0].maketrans('','','.,; ')==ult_camps[0].maketrans('','','.,; ')):
                                            print ('Camps: ',ult_camps,camps)
                                            camps=ult_camps
                                            x=4

                                if (x>3)  : 
                                    ult_camps=camps 
                                    #os.makedirs(save_dir, exist_ok=True)
                                    try:
                                        textdata = camps[2]
                                        if(camps[2].find(r'.')!=-1):
                                            textdata = camps[2].replace(r'.','-')  
                                        if (textdata.find(r',')!=-1):
                                            textdata = textdata.replace(r',','-')  
                                        if (textdata.find(r':')!=-1):
                                            textdata = textdata.replace(r':','-')   
                                        
                                        if(textdata.find(r'/')!=-1):
                                            textdata = textdata.replace(r'/','-')
                                        if (camps[1].find(r'/')!=-1):    
                                            camps[1]=camps[1].replace(r'/','-')
                                        if (camps[3].find(r'/')!=-1):    
                                            camps[3]=camps[3].replace(r'/','-')
                                        try:    
                                            data_obj = datetime.strptime(textdata,'%d-%m-%y')
                                        except:
                                            try:
                                                data_obj = datetime.strptime(textdata,'%d-%m-%Y')
                                            except:   
                                                print('Error en la data ',textdata)
                                                self.guardar_logs('La data no te format correcte',1,2, basename + f"_{page+1}.pdf")
                                                sw_error=True   
                                    except:
                                        sw_error=True
                                    
                                                                        
                                    if (sw_error==False):
                                        carpeta=self.gestiona_dirs(save_dir,plataforma,data_obj,plant)
        
                                        ClientNav=None
                                        # print('*************'+plataforma+'*************'+plant+'*************'+camps[0]+'*************')
                                        ClientNav=self.selectCodiClientNav(plataforma,plant,camps[0])
                                        gNeg=''
                                        if (ClientNav is not None):
                                            if(ClientNav[3] >' '):
                                                CodiClientNav=ClientNav[3]
                                            else:
                                                CodiClientNav=camps[0]
                                            
                                            if(ClientNav[4] is not None):    
                                                gNeg=ClientNav[4]
                                        else:
                                            CodiClientNav=camps[0]

                                        nom = os.path.join(carpeta,self.gestionanom(CodiClientNav,gNeg, data_obj.strftime("%Y-%m-%d"), plant,camps[3]+".pdf"))
                                else :
                                    #print ('Menys de 4 camps')
                                    self.guardar_logs('No s\'han trobat tots els camps',1,2, basename + f"_{page+1}.pdf")
                                    sw_error=True
                            else: 
                                print ('No s\'ha trobat la plantilla')
                                self.guardar_logs('No s\'ha trobat la plantilla',1,2, basename + f"_{page+1}.pdf")
                                sw_error=True

                        if (sw_error == True):
                                self.revisions=self.revisions+1
                                carpeta=self.gestiona_dirs_errors(save_dir,plataforma,'revisions')
                                nomdades=os.path.join(carpeta,datetime.now().strftime("%Y-%m-%d")+'_'+basename + f"_{page+1}.json")
                                with open(nomdades, "w", encoding="utf-8") as fitxer:
                                                  json.dump(self.jsondades, fitxer, ensure_ascii=False, indent=4)
                                if(sw_guardarjsonerronis or sw_guardarjson):
                                    self.save_json_result( page_res, os.path.join(save_dir,'jsons', basename + f"_{page+1}.json"))
                                ClientNav=None
                                nom=os.path.join(carpeta,datetime.now().strftime("%Y-%m-%d")+'_'+basename + f"_{page+1}.pdf")
                        else :
                            self.correctes=self.correctes+1  
                                 
                        #print(nom)     
                        #print('Guarda document '+nom+' '+datetime.now().strftime("%Y-%m-%d %H:%M:%S")) 
                        self.guardarPDF(nom,img,carpeta)
                        self.guardar_logs('Processat '+plataforma+' ('+basename + f"_{page+1}.pdf"+')',0,1,nom,None,self.reglog)
                        #self.reglog+=nom
                        ############################################ Modul Enviament Emails #############################
                        if(plant=='VP'):
                            if (ClientNav is not None):  #nomes als albarans propis de distribució
                                #print(ClientNav)
                                numemails=len(ClientNav)
                                if (ClientNav[7] >' ') and (not sw_error):
                                    emails=ClientNav[7]
                                    if (numemails>8):
                                        emails=emails+','+ClientNav[8]
                                        if (numemails>9):
                                            emails=emails +','+ClientNav[9]

                                    Envio=Email(emails,"Albarà/Albaran Serhs Distribució",f"{ClientNav[6]}", f"{camps[3]}" ,nom)
                                    Envio.send()
                                    self.guardar_logs(f"- email enviat en {ClientNav[6]} per a "+emails,0,1,nom)

                        ############################################ Modul Enviament Emails #############################
                        self.llistplanes.append(nom)
                        #self.reglog+='\r\n'        
                        # print(self.reglog)  
                        # print('Guarda log Plataforma '+basename+' '+datetime.now().strftime("%Y-%m-%d %H:%M:%S"))          
                        #self.save_log()   
                        if visualize:
                            #self.visualize_image(img, page_res, nom.replace(".pdf",".jpg"))
                            if (sw_error==True):
                                try:
                                    self.visualize_image(img, page_res, nom.replace(".pdf",".jpg"))
                                except:
                                    pass
                    res_list.append(pdf_res)
                    # print('Elimina document '+basename+' '+datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
                    os.remove(fpath)  
                          
            if(self.numPlanes>0): 
                t_estadistica=('📄'+str(self.numPlanes)+'#💾'+str(self.documents)+'#🟢'+str(self.correctes)+'#🟡'+str(self.revisions)+'#❌'+str(self.eliminats))
                self.guardar_logs('Procés finalitzat',0,0,self.numPlanes,t_estadistica)
        
        if (sw_drivepujada):
            if(os.path.isdir('/srv/pdf-extract')):
                # DADES PEL SERVIDOR DE PRODUCCIO  ############################################################################
                run(["rclone", "move", "/srv/pdf-extract/PDF-Extract-Kit-main/outputs/ocr/Palafolls/2025", "ocr_output_palafolls:"])
                run(["rclone", "move", "/srv/pdf-extract/PDF-Extract-Kit-main/outputs/ocr/Tarragona/2025", "ocr_output_tarragona:"])
                run(["rclone", "move", "/srv/pdf-extract/PDF-Extract-Kit-main/outputs/ocr/Ripollet/2025", "ocr_output_ripollet:"])                
                run(["rclone", "move", "/srv/pdf-extract/PDF-Extract-Kit-main/outputs/ocr/Fornells/2025", "ocr_output_fornells:"])
                run(["rclone", "move", "/srv/pdf-extract/PDF-Extract-Kit-main/outputs/ocr/Palafolls/revisions", "ocr_output_palafolls:/revisions"]) 
                run(["rclone", "move", "/srv/pdf-extract/PDF-Extract-Kit-main/outputs/ocr/Tarragona/revisions", "ocr_output_tarragona:/revisions"])
                run(["rclone", "move", "/srv/pdf-extract/PDF-Extract-Kit-main/outputs/ocr/Ripollet/revisions", "ocr_output_ripollet:/revisions"])
                run(["rclone", "move", "/srv/pdf-extract/PDF-Extract-Kit-main/outputs/ocr/Fornells/revisions", "ocr_output_fornells:/revisions"])
            else :
                if (os.path.isdir(r'c:\PDF-Extract-Kit-main\PDF-Extract-Kit-main')): 
                    # DADES PER L'EXECUCIO EN LOCAL  ##############################################################################
                    run(["rclone", "move", r"C:\PDF-Extract-Kit-main\PDF-Extract-Kit-main\outputs\ocr\Palafolls\2025", "ocr_output_palafolls:"])
                    run(["rclone", "move", r"C:\PDF-Extract-Kit-main\PDF-Extract-Kit-main\outputs\ocr\Tarragona\2025", "ocr_output_tarragona:"])
                    run(["rclone", "move", r"c:\PDF-Extract-Kit-main\PDF-Extract-Kit-main\outputs\ocr\Ripollet\2025",  "ocr_output_ripollet:"])
                    run(["rclone", "move", r"c:\PDF-Extract-Kit-main\PDF-Extract-Kit-main\outputs\ocr\Fornells\2025",  "ocr_output_fornells:"])
                    run(["rclone", "move", r"c:\PDF-Extract-Kit-main\PDF-Extract-Kit-main\outputs\ocr\Palafolls\revisions", "ocr_output_palafolls:/revisions"])
                    run(["rclone", "move", r"c:\PDF-Extract-Kit-main\PDF-Extract-Kit-main\outputs\ocr\Tarragona\revisions", "ocr_output_tarragona:/revisions"])
                    run(["rclone", "move", r"c:\PDF-Extract-Kit-main\PDF-Extract-Kit-main\outputs\ocr\Ripollet\revisions",  "ocr_output_ripollet:/revisions"])
                    run(["rclone", "move", r"C:\PDF-Extract-Kit-main\PDF-Extract-Kit-main\outputs\ocr\Fornells\revisions",  "ocr_output_fornells:/revisions"])
                else:
                    # DADES PER L'EXECUCIO EN LOCAL  ##############################################################################
                    run(["rclone", "move", "/mnt/c/PDF-Extract-Kit-main/PDF-Extract-Kit-main/outputs/ocr/Palafolls/2025", "ocr_output_palafolls:"])
                    run(["rclone", "move", "/mnt/c/PDF-Extract-Kit-main/PDF-Extract-Kit-main/outputs/ocr/Tarragona/2025", "ocr_output_tarragona:"])
                    run(["rclone", "move", "/mnt/c/PDF-Extract-Kit-main/PDF-Extract-Kit-main/outputs/ocr/Ripollet/2025",  "ocr_output_ripollet:"])
                    run(["rclone", "move", "/mnt/c/PDF-Extract-Kit-main/PDF-Extract-Kit-main/outputs/ocr/Fornells/2025",  "ocr_output_fornells:"])
                    run(["rclone", "move", "/mnt/c/PDF-Extract-Kit-main/PDF-Extract-Kit-main/outputs/ocr/Palafolls/revisions", "ocr_output_palafolls:/revisions"])
                    run(["rclone", "move", "/mnt/c/PDF-Extract-Kit-main/PDF-Extract-Kit-main/outputs/ocr/Tarragona/revisions", "ocr_output_tarragona:/revisions"])
                    run(["rclone", "move", "/mnt/c/PDF-Extract-Kit-main/PDF-Extract-Kit-main/outputs/ocr/Ripollet/revisions",  "ocr_output_ripollet:/revisions"])
                    run(["rclone", "move", "/mnt/c/PDF-Extract-Kit-main/PDF-Extract-Kit-main/outputs/ocr/Fornells/revisions",  "ocr_output_fornells:/revisions"])
        return res_list
    
    def visualize_image(self, imatge, ocr_res, save_path, cate2color={}):
        """plot each result's bbox and category on image.
        
        Args:
            image: PIL.Image.Image
            ocr_res: list of ocr det and rec, whose format following the results of self.predict_image function
            save_path: path to save visualized image
        """
        try:
            draw = ImageDraw.Draw(imatge)
        except Exception as e:
            print (f"No s\'ha pogut carregar la imatge : Excepció {e}")
            return
        for res in ocr_res:
            #print(res['poly'][0], res['poly'][1], res['poly'][2], res['poly'][3], res['poly'][4], res['poly'][5],res['poly'][6], res['poly'][7], res['text'])
           # box_color = cate2color.get(res['category_type'], (255, 0, 0))
            x_min, y_min = int(res['poly'][0]), int(res['poly'][1])
            x_max, y_max = int(res['poly'][4]), int(res['poly'][5])
            #draw.rectangle([x_min, y_min, x_max, y_max], fill=(205, 245, 186), outline=(255, 0, 0), width=1) 
            if(x_min<x_max and y_min<y_max):
                draw.rectangle([x_min, y_min, x_max, y_max], outline=(255, 0, 0), width=1) 
            else: 
                print('Error de coordenades :',x_min, y_min, x_max, y_max, res['text'])
            #draw.text ((x_min+10, y_min+5), res['text'], (0, 0, 0))
            #draw.text(((x_min+20), (y_min-5)),  res['poly'][0], (0, 255, 0))
        for res in self.cuadresplantilla:
            x_min, y_min = int(res['coord'][0])+ self.marge_document_x , int(res['coord'][1])+ self.marge_document_y 
            x_max, y_max = int(res['coord'][4])+ self.marge_document_x , int(res['coord'][5])+ self.marge_document_y 
            draw.rectangle([x_min, y_min, x_max, y_max], fill=None, outline=(0, 0, 255), width=1)        
            #draw.text ((x_min, y_min-10), res['nom'], (0, 0, 255))
        if save_path:
            imatge.save(save_path)
        #print(self.cuadresplantilla)
        
    def save_json_result(self, ocr_res, save_path):
        """save results to a json file.
        
        Args:
            ocr_res: list of ocr det and rec, whose format following the results of self.predict_image function
            save_path: path to save visualized image
        """
        with open(save_path, "w", encoding="utf-8") as f:
            f.write(json.dumps(ocr_res, indent=2, ensure_ascii=False))
   
    def detectarPlantilla(self, ocr_res):
        nom="assets/inputs/plantilles SDL/plantilla.txt"
        textDocument=json.dumps(ocr_res)
        self.jsondades=[]
        #print(textDocument)
        try:
            with open(nom, 'r', encoding='utf-8') as file:
                data = file.read()
                for linea in data.split("\n"):
                    if linea is not None :
                        arrlinea=linea.split("=")
                    # print('Prova plantilla :',arrlinea)
                        for i in ocr_res:
                            # if (i['score']<self.score_valid):
                            #     continue
                            arrpattern = arrlinea[2].split("|")
                            for p in arrpattern:
                                if (re.search(p,i['text']) is not None):
                                    self.marge_document_y=float(i['poly'][1])-float(arrlinea[3])
                                    self.marge_document_x=float(i['poly'][0])-float(arrlinea[4])
                                #  print ('Limits plantilla :',i['poly'][0],' ',i['poly'][1], '   ', arrlinea[4], ' ', arrlinea[3])
                                    return [arrlinea[0],arrlinea[1]]
                            # if (re.search(arrlinea[2],i['text']) is not None):    
                            #     self.marge_document_y=float(i['poly'][1])-float(arrlinea[3])
                            #     self.marge_document_x=float(i['poly'][0])-float(arrlinea[4])
                            #     return [arrlinea[0],arrlinea[1]]
 
        except FileNotFoundError:
            print("Error: l'arxiu "+ nom +" no hi és.")
        return []
    
    def detectarCamps(self, plant, ocr_res):
        nom="assets/inputs/plantilles SDL/"+plant+".txt"
        resultat=[]
        cont=-1
        self.cuadresplantilla=[]
        self.jsondades=[]
        registrejson='"plantilla": "'+plant+'",'
        try:
            with open(nom, 'r', encoding='utf-8') as file:
                data = file.read()
                for linea in data.split("\n"):
                    if (linea[0]=="/") and (linea[1]=="*"):  # ignora comentaris
                        continue
                    cont=cont+1
                    arrlinea=linea.split("=")
                    camp=arrlinea[0]
                    arrposicio=arrlinea[1].split(",") 
                    expresio=arrlinea[2]
                    #print (f'Expressio : {camp} :',expresio)
                    posicio=[]
                    for a in arrposicio:
                        posicio.append(float(a))
                    self.cuadresplantilla.append({'nom':camp,'coord':posicio})
                    #print(posicio)
                    for i in ocr_res:
                        if (i['score']<self.score_valid):
                            continue
                        #print (' Camp : '+ camp )
                        poly = []
                        for el in i['poly']:
                            poly.append(float(el))
                        #if(i['poly']==posicio):
                        if(abs(i['poly'][1]-(posicio[1]+self.marge_document_y))<self.marge_posicio_y):
                            v_x=self.marge_posicio_x + abs(self.marge_document_x)    
                            if(abs(i['poly'][0]-posicio[0])+self.marge_document_x <self.marge_posicio_x) or ((i['poly'][0]<(posicio[0]+v_x) or i['poly'][0]<(posicio[0]-v_x)) and (i['poly'][2]>(posicio[2]+v_x ) or i['poly'][2]>(posicio[2]-v_x ))):
                        #if (i['poly'][0]<(posicio[0]+v_x) and i['poly'][0]>(posicio[0]-v_x) and i['poly'][2]<(posicio[2]+v_x ) and i['poly'][2]>(posicio[2]-v_x )): # <--- Falta perfeccionar aquesta condicio 
                        #if(i['poly'][0] <(posicio[0]+self.marge_posicio)): 
                            #print('marge total',v_x)  
                                                       # print('posicion Plantilla  :',posicio[0],posicio[1],posicio[2],posicio[3],camp,self.marge_document_x,self.marge_document_y)
                                                       # print('posicion Camp OCR   :',i['poly'][0],i['poly'][1],i['poly'][2],i['poly'][3],i['text'],self.marge_document_x,self.marge_document_y)
                                                        registrejson+='"'+camp.lower()+'":"'+i['text']+'",'
                                                            
                                    # if (i['poly'][0]<(posicio[0]) and i['poly'][2]>(posicio[2])): 
                                    #                     print('Includes posicio   :',posicio[0],posicio[1],posicio[2],posicio[3],camp)
                                    #                     print('Includes Camp OCR   :',i['poly'][0],i['poly'][1],i['poly'][2],i['poly'][3],i['text'])    
                               #if (abs(i['poly'][2]-posicio[2])<self.marge_posicio ): # allargada del camp no ho te en compte
                                   #if (abs(i['poly'][3]-posicio[3])<self.marge_posicio+self.marge_document):
                                        #if (abs(i['poly'][4]-posicio[4])<self.marge_posicio  ): # allargada del camp no ho te en compte
                                          #  if (abs(i['poly'][5]-posicio[5])<self.marge_posicio+self.marge_document):
                                            #    if (abs(i['poly'][6]-posicio[6])<self.marge_posicio+self.marge_document):
                                                  #  if (abs(i['poly'][7]-posicio[7])<self.marge_posicio+self.marge_document):
                                                       # print ('El camp '+camp+' es: ',i['text'])
                                                        if (expresio>' '):
                                                            # arrpattern = expresio.split("|")
                                                            # variable=None
                                                            # for p in arrpattern:
                                                            #     variable=re.search(p, i['text'])
                                                            #     if (variable is not None):
                                                            #         break
                                                            variable=re.search(expresio, i['text'])
                                                            
                                                            if(variable is not None):
                                                                #print ('El camp '+camp+' es: ',variable.group(0))
                                                                resultat.append(variable.group(0))
                                                            # else:
                                                            #     print ('No TROBAT el camp '+camp+' : ',i['text'])
                                                        else:
                                                            resultat.append(i['text'])        
                #print (resultat)
                registrejson=registrejson[:-1]
                self.jsondades='{'+registrejson+'}' 
                print ('json:',self.jsondades)

                
                self.reglog = str(resultat)
        except FileNotFoundError:
            print("Error: l'arxiu "+ nom +" no hi és.")
        return resultat
    def carregaDadesNavision(self):
        with open('assets/inputs/CliGrupNegoci.csv', encoding='latin1') as f:
          cont=0
          for linea in f:
            cont=cont+1
            self.dadesGrupNegoci.append(linea.strip().split(';')) 
        print('Clients carregats :',cont)    

    def selectCodiClientNav(self, plataforma, prove, codi):
        if (prove!='CCEP' and prove != 'DDI'): prove='*'
        for i in self.dadesGrupNegoci:
            if(codi==i[1]):
                if(prove==i[0] or prove=='*'):
                    if(prove=='*' and ('SERHS'==i[0] or'VP'==i[0])):
                        if(plataforma==i[2]):
                            #print(datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
                            return i
                    else:       
                        if(plataforma==i[2]):
                            #print(datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
                            return i
        return None

    # def save_log(self):
    #     with open(self.logfilename, 'a', encoding='utf-8') as self.log:
    #         self.log.write(self.reglog)
    #         self.reglog=''

    def gestiona_dirs(self,save_dir,dir,data_obj,plant):
        if not os.path.exists(save_dir):
            os.makedirs(save_dir)
        if not os.path.exists(os.path.join(save_dir,dir)):
            os.makedirs(os.path.join(save_dir,dir))
        if not os.path.exists(os.path.join(save_dir,dir,str(data_obj.year))):
            os.makedirs(os.path.join(save_dir, dir, str(data_obj.year)))
        if not os.path.exists(os.path.join(save_dir,dir, str(data_obj.year), str(data_obj.month))):
            os.makedirs(os.path.join(save_dir, dir, str(data_obj.year), str(data_obj.month)))
        if not os.path.exists(os.path.join(save_dir, dir, str(data_obj.year), str(data_obj.month), str(data_obj.day))):
            os.makedirs(os.path.join(save_dir, dir, str(data_obj.year), str(data_obj.month), str(data_obj.day)))
        if not os.path.exists(os.path.join(save_dir, dir,str(data_obj.year), str(data_obj.month), str(data_obj.day), plant)):
            os.makedirs(os.path.join(save_dir, dir,str(data_obj.year), str(data_obj.month), str(data_obj.day), plant))
        return os.path.join(save_dir, dir, str(data_obj.year), str(data_obj.month), str(data_obj.day), plant)
       
    def gestiona_dirs_errors(self,save_dir,dir,plant):
        if not os.path.exists(save_dir):
            os.makedirs(save_dir)
        if not os.path.exists(os.path.join(save_dir,dir)):
            os.makedirs(os.path.join(save_dir,dir))
        if not os.path.exists(os.path.join(save_dir, dir, plant)):
            os.makedirs(os.path.join(save_dir, dir, plant))
        return os.path.join(save_dir, dir,  plant)
   
    def gestionanom(selft,client, gNeg, dataymd, plant,comanda):
        nom = client + '_' 
        if (gNeg is not None) and (gNeg > ' '):
            nom += '[' + gNeg + ']_'
        nom += dataymd +'_'  + plant + '_' + comanda
        return nom 

    def guardarPDF(self,nom, img, carpeta):
        if nom in self.llistplanes:
            nom2 = carpeta + "/temporal.pdf"
            #print('Nom 2 ' ,nom2)
            img.save(nom2)
            self.afegir_pdfs(nom,nom2)
        else:
            # print('Nom únic',nom)
            img.save(nom)
            self.documents=self.documents+1
            
    def afegir_pdfs(self,pdf_base, pdf_nou):
        sortida=None
        writer = PdfWriter()
        cont=0
        for path in [pdf_base, pdf_nou]:
            reader = PdfReader(path)
            for page in reader.pages:
                writer.add_page(page)
                cont=cont+1
        self.reglog += ' Plana '+str(cont)+' afegida '
        if sortida is None:
            sortida = pdf_base  # sobreescriu el base

        with open(sortida, "wb") as f:
            writer.write(f)
            
        if os.path.exists(pdf_nou):
            os.remove(pdf_nou)
            
    # def separa_planes(self, nom, carpeta):
    #     print (nom)
    #     reader = PdfReader(nom)
    #     print ('Planes ',len(reader.pages))
    #     for i, page in enumerate(reader.pages, start=1):
    #         writer = PdfWriter()
    #         writer.add_page(page)
    #         nomsenseext=nom.replace('.pdf','')
    #         with open(nomsenseext + f"__{i:03d}.pdf", "wb") as f:
    #             writer.write(f)
    #     os.remove(nom)

    def detect_orientation(selft,image):
        #Detecta l’orientació mitjançant Tesseract (OSD).
        try:
            osd = pytesseract.image_to_osd(image, output_type=pytesseract.Output.DICT)
            return osd.get("rotate", 0)
        except pytesseract.TesseractError as e:
            print(f"⚠️  Error d’OSD: {e}. S’assumeix 0°.")
            return 0

    def separa_i_orienta(self,input_pdf):
        #Corregeix l’orientació de cada pàgina i la guarda en un fitxer individual.
        #Els fitxers tindran el nom: {output_prefix}_1.pdf, {output_prefix}_2.pdf, etc.
        # if "_Pag_" in input_pdf:
        #     return 0
        
        reader = PdfReader(input_pdf)
        doc = fitz.open(input_pdf)
        nomsenseext=input_pdf.replace('.pdf','')
        total_pages = len(reader.pages)
        #if(total_pages>1):
        print(f"📄 Processant {total_pages} pàgines...")
        if(total_pages>1):
                sw_guardar=True
        else :
                sw_guardar=False
        for page_number in range(total_pages):
            pdf_page = reader.pages[page_number]
            fitz_page = doc[page_number]

            # Renderitzar com a imatge
            pix = fitz_page.get_pixmap(dpi=150)
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            if ('r-00' in input_pdf.lower()):
                print(f"✓ Pàgina {page_number + 1}: orientació correcta")
            else :
             if ('r90_' in input_pdf.lower()):
                    pdf_page.rotate(90)
                    print(f"↻ Pàgina {page_number + 1}: rotació assignada 90° → corregint")
                    sw_guardar=True    
             else :
                if ('r-90' in input_pdf.lower()):
                    pdf_page.rotate(-90)
                    print(f"↻ Pàgina {page_number + 1}: rotació assignada -90° → corregint")
                    sw_guardar=True
                else:
                    if ('r-180' in input_pdf.lower()):
                        pdf_page.rotate(180)
                        print(f"↻ Pàgina {page_number + 1}: rotació assignada 180° → corregint")
                        sw_guardar=True
                    else:
                        # Detectar i corregir orientació
                        angle = self.detect_orientation(img)
                        if angle != 0:
                            print(f"↻ Pàgina {page_number + 1}: rotació detectada {angle}° → corregint")
                            pdf_page.rotate(angle)
                            sw_guardar=True
                        else:
                            print(f"✓ Pàgina {page_number + 1}: orientació correcta")
                            # pdf_page.rotate(angle)

            # Crear fitxer individual
            if(sw_guardar==True):    
                output_file = nomsenseext
                output_file = re.sub(r'r-90',  'r-00',  output_file)  # cas 1
                output_file = re.sub(r'r-180', 'r-000', output_file)  # cas 2
                output_file = re.sub(r'r90',   'r-00',  output_file)  # cas 3
                print (output_file)
                if(total_pages>1):
                    output_file = f"{output_file}_Pag_{page_number + + 1:04d}.pdf"
                else:
                    output_file = f"{output_file}.pdf"

                writer = PdfWriter()
                writer.add_page(pdf_page)
                with open(output_file, "wb") as f:
                    writer.write(f)
                print(f"   → Guardat: {output_file}")

        doc.close()
        if (total_pages>1):
            print(f"✅ Totes les pàgines separades amb prefix: Pag_*.pdf")
        if sw_guardar:
            if (input_pdf!=output_file):
                os.remove(input_pdf)
        return total_pages

    def guardar_logs(self,descripcio,gravetat,nivell,document,estadistica=None,camps=None):       
        wdata = datetime.now()
        dataprocess= datetime.now().strftime("%Y-%m-%d")
        # [proces][servidor][bdd][nivell][missatge][estadistica][gravetat][timestamp][usuari][text1][text2][text3][text4][text5][text6][text7][text8][text9][text10]
        buffer="'PDF-EXTRACT','10.252.252.32','pdf_extract','"+str(nivell)+"','"+str(descripcio)+"','"+str(estadistica)+"','"+str(gravetat)+"','"+str(wdata)+"',INTNAVANT,'"+str(document)+"','"+str(camps)+"','','','','','','','','',''\n"  #+'PDF-EXTRACT,10.252.252.32,pdf_extract,'+nivell+','+descripcio+','+estadistica+','+gravetat+','+wdata+','+INTNAVANT+','+document+',,,,,,,,,\n'  
        with open('outputs/LogPDF_Extract_'+dataprocess+'.txt', 'a', encoding='utf-8') as logGen:
            logGen.write(buffer)
        print(buffer)

    #def guardar_logsToDB(self,descripcio,gravetat,nivell,document,estadistica=None):
        # conn_str = f'DRIVER={{ODBC Driver 18 for SQL Server}};SERVER=''10.70.1.217'';DATABASE=''TIC-INTERN'';UID=''INTNAVANT'';PWD=''SERHS2015@'';Encrypt=no;'
        # conn = pyodbc.connect(conn_str)
        # cursor = conn.cursor()
        # cursor.execute( """EXEC dbo.sp_InsertarLog
        # @servidor = ?,
        # @bdd = ?,
        # @proces = ?,
        # @nivell = ?,
        # @missatge = ?,
        # @gravetat = ?,
        # @text1 = ?,
        # @estadistica = ?,
        # @Estat = ? """
        # , ('10.252.252.32', 'pdf_extract', 'PDF-EXTRACT', nivell, descripcio, gravetat, document,estadistica,''))

        # #cursor.execute(query_insertarlog)
        # conn.commit()
        # cursor.close()
        # conn.close()
        