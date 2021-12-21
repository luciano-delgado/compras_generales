import openpyxl, tkinter as tk, numpy as np, pandas as pd
from getpass import getuser
from datetime import datetime
from tkinter import *
from tkinter import ttk
from tkinter import messagebox
from enviar_mail import enviarmail

def generar_texto(usu,fac,pro,fh):
    
    array_fact = np.array(fac)
    array_pro = np.array(pro)
    array_fh = np.array(fh)

    df=pd.DataFrame({'Factura':array_fact,'Fecha emision':array_fh,'Proveedor':array_pro,})

    # df['Factura'].str.center(width = 15, fillchar = '%')
    # df['Proveedor'].str.center(width = 15, fillchar = '%')
    # df['Fecha emision'].str.center(width = 15, fillchar = '%')
    
    df2 = df
    cuerpo_mail = f'Estimado/a {usu}, \n\nNecesitamos por favor que generen la recepción y/o autorizacion de las siguientes facturas: \n\n {df2} \n\n\n Muchas gracias.'
    #print(cuerpo_mail)

    return cuerpo_mail
# -----------------------------------------------------------------

#generar_texto('Luciano',['123','2432','435','2777432'],['ACDC','PROV','JCM','KC'],['11.12','09.12','01.12','10.12'])

