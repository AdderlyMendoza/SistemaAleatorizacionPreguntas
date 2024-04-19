from docxtpl import DocxTemplate
import pandas as pd

doc = DocxTemplate("E:\SISTEMA ALEATORIZACION EXAMEN\SISTEMA\plantilla.docx")

df = pd.read_excel("SISTEMA\prueba_ALEATORIZADO.xlsx")

constantes = {}
for i, f in df.iterrows():
    constantes[f'pregunta_{i}'] = df.iloc[i]["ENUNCIADO"]
    constantes[f'A_{i}'] = df.iloc[i]["A"]
    constantes[f'B_{i}'] = df.iloc[i]["B"]
    constantes[f'C_{i}'] = df.iloc[i]["C"]
    constantes[f'D_{i}'] = df.iloc[i]["D"]
    constantes[f'E_{i}'] = df.iloc[i]["E"]

# print(constantes)
    
doc.render(constantes)
doc.save(f"prueba.docx")
print("EXPORTADO!")
