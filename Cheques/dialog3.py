from cheques import buscarVentana
from tkinter import Tk, OptionMenu, Button, StringVar, mainloop, Listbox, MULTIPLE
from datetime import date, timedelta
import logging

OPTIONS = [
"Jan",
"Feb",
"Mar"
] #etc

empresas = [
    ('ACHEGEO', '1102002', 'ACHEGEO.png'),
    ('ANDES', '1102002', 'ANDES.png'),
    ('APOCE', '1102002', 'APOCE.png'),
    ('AUTO PARQUIME', '1102000', 'AUTO PARQUIME.png'),
    ('AUTOORDEN', '1102008', 'AUTOORDEN.png'),
    ('CAFETAL', '1102001', 'CAFETAL.png'),
    ('CAJA EXPRESS', '1102005', 'CAJA EXPRESS.png'),
    ('CLEVER PARK SPA', '1102002', 'CLEVER PARK SPA.png'),
    ('CONCESIONES PROVIDENCIA', '5111042', 'CONCESIONES PROVIDENCIA.png'),
    ('CONCESIONES PUNTA ARENAS', '1102004', 'CONCESIONES PUNTA ARENAS.png'),
    ('CONCESIONES SANTIAGO', '1102003', 'CONCESIONES SANTIAGO.png'),
    ('CONCESIONES RECOLETA', '1102002', 'CONCESIONES RECOLETA.png'),
    ('CONSORCIO VALPARAISO', '1102004', 'CONSORCIO VALPARAISO.png'),
    ('COQUIMBO', '1102001', 'COQUIMBO.png'),
    ('CP LATINA CHILE', '1102000', 'CP LATINA CHILE.png'),
    ('DENSITOSEA', '1102004', 'DENSITOSEA.png'),
    ('DIAGNOSIS', '1102000', 'DIAGNOSIS.png'),
    ('ECM GEOTERM1IA', '1102000', 'ECM GEOTERMIA.png'),
    ('ESTACIONAR', '1102003', 'ESTACIONAR.png'),
    ('IMAGEN', '1102000', 'IMAGEN.png'),
    ('INFINERGEO', '1110132', 'INFINERGEO.png'),
    ('INGENIEROS', '1102003', 'INGENIEROS.png'),
    ('INVER EST', '1102002', 'INVER EST.png'),
    ('INVERREST', '1102008', 'INVERREST.png'),
    ('INVERTERRA', '1102004', 'INVERTERRA.png'),
    ('IQUIQUE', '1102004', 'IQUIQUE.png'),
    ('MEDCONSUL', '1102003', 'MEDCONSUL.png'),
    ('PIEDMONT', '1102005', 'PIEDMONT.png'),
    ('POLLO ABRAZADO', '1102001', 'POLLO ABRAZADO.png'),
    ('RENAL', '1102003', 'RENAL.png'),
    ('SEC', '1102004', 'SEC.png'),
    ('SERVILAND', '1102008', 'SERVILAND.png'),
    ('TALCA', '1102002', 'TALCA.png'),
    ('TOMO IMAGEN', '1102005', 'TOMO IMAGEN.png'),
    ('VALDIVIA', '1102001', 'VALDIVIA.png'),
    ('VALSEGUR', '1102005', 'VALSEGUR.png'),
    ('VALSEGUR-ETV', '1102005', 'VALSEGUR-ETV.png')
]

hoy = date.today()
hoy_formato_dia = date.today().strftime("%a")

logging.basicConfig(
    filename='Log_Cheques_' + date.today().strftime("%d-%m-%Y") + '.txt',
    level=logging.INFO,
    format='%(asctime)s.%(msecs)03d: %(message)s',
    datefmt='%H:%M:%S')

if hoy_formato_dia == 'Mon':
    fechaFlex = hoy - timedelta(days=3)
    fechaFlex = fechaFlex.strftime("%d-%m-%Y").replace("-", "")
else:
    fechaFlex = hoy - timedelta(days=1)
    fechaFlex = fechaFlex.strftime("%d-%m-%Y").replace("-", "")

fecha = date.today().strftime("%d-%m-%Y")

print(fecha)

root = Tk()
root.geometry('180x700')

# Create a listbox
listbox = Listbox(root, width=40, height=700, selectmode=MULTIPLE)

# Inserting the listbox items

listbox.insert(1, "ACHEGEO")                                                                                                                                                                                                                            
listbox.insert(2, "ANDES")
listbox.insert(3, "APOCE")                                                                                               
listbox.insert(4, "AUTO PARQUIME")                                                                                     
listbox.insert(5, "AUTOORDEN")                                                                                          
listbox.insert(6, "CAFETAL")                                                                                             
listbox.insert(7, "CAJA EXPRESS")                                                                                       
listbox.insert(8, "CLEVER PARK SPA")                                                                                    
listbox.insert(9, "CONCESIONES PROVIDENCIA")
listbox.insert(10, "CONCESIONES PUNTA ARENAS")                                                                           
listbox.insert(11, "CONCESIONES SANTIAGO")
listbox.insert(12, "CONCESIONES RECOLETA")                                                                              
listbox.insert(13, "CONSORCIO VALPARAISO")                                                                              
listbox.insert(14, "COQUIMBO")                                                                                 
listbox.insert(15, "CP LATINA CHILE")                                                                                   
listbox.insert(16, "DENSITOSEA")                                                                                        
listbox.insert(17, "DIAGNOSIS")                                                                                         
listbox.insert(18, "ECM GEOTERM1IA")                                                                                    
listbox.insert(19, "ESTACIONAR")                                                                                        
listbox.insert(20, "IMAGEN")                                                                                            
listbox.insert(21, "INFINERGEO")
listbox.insert(22, "INGENIEROS")                                                                                        
listbox.insert(23, "INVER EST")
listbox.insert(24, "INVERREST")
listbox.insert(25, "INVERTERRA")                                                                                    
listbox.insert(26, "IQUIQUE")                                                                                           
listbox.insert(27, "MEDCONSUL")                                                                                         
listbox.insert(28, "PIEDMONT")                                                                                          
listbox.insert(29, "POLLO ABRAZADO")                                                                                    
listbox.insert(30, "RENAL")
listbox.insert(31, "SEC")
listbox.insert(32, "SERVILAND")
listbox.insert(33, "TALCA")
listbox.insert(34, "TOMO IMAGEN")                                                                                       
listbox.insert(35, "VALDIVIA")
listbox.insert(36, "VALSEGUR")                                                                                          
listbox.insert(37, "VALSEGUR-ETV")
# Function for printing the
# selected listbox value(s)
def selected_item():
	
	# Traverse the tuple returned by
	# curselection method and print
	# corresponding value(s) in the listbox
	for i in listbox.curselection():
		print(listbox.get(i))

# Create a button widget and
# map the command parameter to
# selected_item function
btn = Button(root, text='Print Selected', command=selected_item)

# Placing the button and listbox
btn.pack(side='bottom')
listbox.pack()

root.mainloop()
