import os
import pandas as pd

excel = os.getcwd() + '\\dev_comp.xlsx'
pathOUT = os.getcwd() + '\\output.xlsx'


def unionGYH(df):
    G = df.iloc[:,6].values.tolist()
    H = df.iloc[:,7].values.tolist()
    GyH = df.iloc[:,6].values.tolist()

    for i in range(len(G)):
        G[i] = str(G[i])
        H[i] = str(H[i])
        while(len(G[i]) < 2):
            G[i] = "0" + G[i]
        while(len(H[i]) < 9):
            H[i] = "0" + H[i]

        GyH[i] = G[i] + H[i]

    return pd.DataFrame(GyH, columns = ["comp"])
    

def llenarColumnaDeStr(df,cadena, titulo = "column"):
    aux = df.iloc[:,0].values.tolist()
    
    for i in range(len(aux)):
        aux[i] = cadena

    return pd.DataFrame(aux, columns = [titulo])


def sumarStringAColumna(cadena,dfColumna, titulo ="column"):
    aux = dfColumna.values.tolist()

    for i in range(len(aux)):
        aux[i] = cadena + str(aux[i])

    return pd.DataFrame(aux, columns = [titulo])

def estandarizarCUIT(dfColumna):
    payer_id_number = dfColumna.values.tolist()
    payer_id_type = dfColumna.values.tolist()

    for i in range(len(payer_id_type)):
        if(len(str(payer_id_type[i])) < 10):
            payer_id_number[i] = ""
            payer_id_type[i] = ""
        else:
            payer_id_type[i] = "CUIT_ARG"
        

    df_payer_id_type = pd.DataFrame(payer_id_type,columns=["payer_id_number"])
    df_payer_id_number = pd.DataFrame(payer_id_number,columns=["payer_id_number"])

    return pd.concat([df_payer_id_type,df_payer_id_number],axis=1)


df = pd.read_excel(excel)


GyH = unionGYH(df)

#Columnas repetidas



#En el documento final excel hay 3 secciones de colores, las separé en codigo simplemente por comodidad de lectura

verde = pd.concat([GyH,GyH,GyH,llenarColumnaDeStr(df,"","concept_description"),llenarColumnaDeStr(df,"ARS","moneda"),df.iloc[:,32], llenarColumnaDeStr(df,"","due_date"), llenarColumnaDeStr(df,"","last_due_date")], axis=1)
azul = pd.concat([llenarColumnaDeStr(df,"","return_url"),llenarColumnaDeStr(df,"","back_url"),llenarColumnaDeStr(df,"","notification_url"),llenarColumnaDeStr(df,"","rate"),llenarColumnaDeStr(df,"","charge_delay"),llenarColumnaDeStr(df,"","payment_number"),llenarColumnaDeStr(df,"","promotion_code"),llenarColumnaDeStr(df,"","meta_data")], axis=1)
amarillo = pd.concat([sumarStringAColumna("",df.iloc[:,3],"payer_reference"),df.iloc[:,17],llenarColumnaDeStr(df,"","payer_email"),llenarColumnaDeStr(df,"","payer_phone"),llenarColumnaDeStr(df,"ARG","payer_id_country"), estandarizarCUIT(df.iloc[:,15]),df.iloc[:,0],df.iloc[:,1]],axis=1)

df = pd.concat([verde,azul,amarillo],axis=1)
#print(df)

df.to_excel(pathOUT, index = False)

#Test

