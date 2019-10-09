#importa as bibliotecas
import xlwt as xml
import xlsxwriter
import matplotlib.pyplot as plt
from xlwt import Workbook
from time import sleep

#inserção de dados do fornecedor no sistema no sistema
print ("CHECKLIST DE HOMOLOGAÇÃO")
sleep(1)

checklist = True

name = input("Nome: ").upper()
cnpj = input("CNPJ: ").upper()
city = input("Localidade: ").upper()
sleep(1)

print("Análise técnica {}, CNPJ {}, localizado em {}.".format(name, cnpj, city))
sleep(1)

print ("Ensaios")

# 1 
f_media_ok = None
while f_media_ok not in ("sim", "não"):
    f_media_ok = input("Pergunta 1")
    if f_media_ok == "sim":
         f_media = input("Pergunta 1: ")
         if f_media == "cubica":
             print ("Informações validadas - APROVADO")
             f_media_r = "APROVADO"
         else:
             print ("Informações validadas - REPROVADO")
             f_media_r = "REPROVADO"
    elif f_media_ok == "nao":
         print ("Informações validadas - ENVIO PENDENTE")
         f_media_r = "ENVIO PENDENTE"
else:
    print("Por favor, preencha sim ou não")
sleep(1)

# 2 
f_cub_ok = None
while f_cub_ok not in ("sim", "não"):
    f_cub_ok = input("Pergunta 2")
    if f_cub_ok == "sim":
         f_cub = float(input("Pergunta 2: "))
         if f_cub <= 15:
             print ("Informações validadas - APROVADO")
             f_cub_r = "APROVADO"
         else:
             print ("Informações validadas - REPROVADO")
             f_cub_r = "REPROVADO"
    elif f_cub_ok == "nao":
         print ("Informações validadas - ENVIO PENDENTE")
         f_cub_r = "ENVIO PENDENTE"
    else:
    	print("Por favor, preencha sim ou não")
sleep(1)

# 3
m_esp_ok = None
while m_esp_ok not in ("sim", "não"):
    m_esp_ok = input("Pergunta 3")
    if m_esp_ok == "sim":
         m_esp = float(input("Pergunta 3"))
         if m_esp >= 2500:
             print ("Informações validadas - APROVADO")
             m_esp_r = "APROVADO"
         else:
             print ("Informações validadas - REPROVADO")
             m_esp_r = "REPROVADO"
    elif m_esp_ok == "nao":
         print ("Informações validadas - ENVIO PENDENTE")
         m_esp_r = "ENVIO PENDENTE"
    else:
    	print("Por favor, preencha sim ou não")
sleep(1)

# 4 
ab_agua_ok = None
while ab_agua_ok not in ("sim", "não"):
    ab_agua_ok = input("Pergunta 4")
    if ab_agua_ok == "sim":
         ab_agua = float(input("Pergunta 4"))
         if ab_agua <= 0.8:
             print ("Informações validadas - APROVADO")
             ab_agua_r = "APROVADO"
         else:
             print ("Informações validadas - REPROVADO")
             ab_agua_r = "REPROVADO"
    elif ab_agua_ok == "nao":
         print ("Informações validadas - ENVIO PENDENTE")
         ab_agua_r = "ENVIO PENDENTE"
    else:
    	print("Por favor, preencha sim ou não")
sleep(1)

# 5 
p_apar_ok = None
while p_apar_ok not in ("sim", "não"):
    p_apar_ok = input("Pergunta 5")
    if p_apar_ok == "sim":
         p_apar = float(input("Pergunta 5"))
         if p_apar <= 15:
             print ("Informações validadas - APROVADO")
             p_apar_r = "APROVADO"
         else:
             print ("Informações validadas - REPROVADO")
             p_apar_r = "REPROVADO"
    elif p_apar_ok == "nao":
         print ("Informações validadas - ENVIO PENDENTE")
         p_apar_r = "ENVIO PENDENTE"
    else:
    	print("Por favor, preencha sim ou não")
sleep(1)

# 6 
r_inter_ok = None
while r_inter_ok not in ("sim", "não"):
    r_inter_ok = input("Pergunta 6")
    if r_inter_ok == "sim":
         r_inter = float(input("Pergunta 6"))
         if r_inter <= 10:
             print ("Informações validadas - APROVADO")
             r_inter_r = "APROVADO"
         else:
             print ("Informações validadas - REPROVADO")
             r_inter_r = "REPROVADO"
    elif r_inter_ok == "nao":
         print ("Informações validadas - ENVIO PENDENTE")
         r_inter_r = "ENVIO PENDENTE"
    else:
    	print("Por favor, preencha sim ou não")
sleep(1)

# 7 
r_comp_ok = None
while r_comp_ok not in ("sim", "não"):
    r_comp_ok = input("Pergunta 7")
    if r_comp_ok == "sim":
         r_comp = float(input("Pergunta 7"))
         if r_comp >= 100:
             print ("Informações validadas - APROVADO")
             r_comp_r = "APROVADO"
         else:
             print ("Informações validadas - REPROVADO")
             r_comp_r = "REPROVADO"
    elif r_comp_ok == "nao":
         print ("Informações validadas - ENVIO PENDENTE")
         r_comp_r = "ENVIO PENDENTE"
    else:
    	print("Por favor, preencha sim ou não")
sleep(1)

# 8
r_choque_ok = None
while r_choque_ok not in ("sim", "não"):
    r_choque_ok = input("Pergunta 8")
    if r_choque_ok == "sim":
         r_choque = float(input("Pergunta 8"))
         if r_choque <= 25:
             print ("Informações validadas - APROVADO")
             r_choque_r = "APROVADO"
         else:
             print ("Informações validadas - REPROVADO")
             r_choque_r = "REPROVADO"
    elif r_choque_ok == "nao":
         print ("Informações validadas - ENVIO PENDENTE")
         r_choque_r = "ENVIO PENDENTE"
    else:
    	print("Por favor, preencha sim ou não")
sleep(1)

# 9 
t_frag_ok = None
while t_frag_ok not in ("sim", "não"):
    t_frag_ok = input("Pergunta 9")
    if t_frag_ok == "sim":
         t_frag = float(input("Pergunta 9"))
         if t_frag <= 5:
             print ("Informações validadas - APROVADO")
             t_frag_r = "APROVADO"
         else:
             print ("Informações validadas - REPROVADO")
             t_frag_r = "REPROVADO"
    elif t_frag_ok == "nao":
         print ("Informações validadas - ENVIO PENDENTE")
         t_frag_r = "ENVIO PENDENTE"
    else:
    	print("Por favor, preencha sim ou não")
sleep(1)

# 10
m_pulv_ok = None
while m_pulv_ok not in ("sim", "não"):
    m_pulv_ok = input("Pergunta 10")
    if m_pulv_ok == "sim":
         m_pulv = float(input("Pergunta 10"))
         if m_pulv <= 1:
             print ("Informações validadas - APROVADO")
             m_pulv_r = "APROVADO"
         else:
             print ("Informações validadas - REPROVADO")
             m_pulv_r = "REPROVADO"
    elif m_pulv_ok == "nao":
         print ("Informações validadas - ENVIO PENDENTE")
         m_pulv_r = "ENVIO PENDENTE"
    else:
    	print("Por favor, preencha sim ou não")
sleep(1)

# 11
t_arg_ok = None
while t_arg_ok not in ("sim", "não"):
    t_arg_ok = input("Pergunta 11")
    if t_arg_ok == "sim":
         t_arg = float(input("Pergunta 11"))
         if t_arg <= 0.5:
             print ("Informações validadas - APROVADO")
             t_arg_r = "APROVADO"
         else:
             print ("Informações validadas - REPROVADO")
             t_arg_r = "REPROVADO"
    elif t_arg_ok == "nao":
         print ("Informações validadas - ENVIO PENDENTE")
         t_arg_r = "ENVIO PENDENTE"
    else:
    	print("Por favor, preencha sim ou não")
sleep(1)

# 12 
m_uni_ok = None
while m_uni_ok not in ("sim", "não"):
    m_uni_ok = input("Pergunta 12")
    if m_uni_ok == "sim":
         m_uni = float(input("Pergunta 12"))
         if m_uni >= 1.25:
             print ("Informações validadas - APROVADO")
             m_uni_r = "APROVADO"
         else:
             print ("Informações validadas - REPROVADO")
             m_uni_r = "REPROVADO"
    elif m_uni_ok == "nao":
         print ("Informações validadas - ENVIO PENDENTE")
         m_uni_r = "ENVIO PENDENTE"
    else:
    	print("Por favor, preencha sim ou não")
sleep(1)

# 13
r_des_ok = None
while r_des_ok not in ("sim", "não"):
    r_des_ok = input("Pergunta 13")
    if r_des_ok == "sim":
         r_des = float(input("Pergunta 13"))
         if r_des <= 30:
             print ("Informações validadas - APROVADO")
             r_des_r = "APROVADO"
         else:
             print ("Informações validadas - REPROVADO")
             r_des_r = "REPROVADO"
    elif r_des_ok == "nao":
         print ("Informações validadas - ENVIO PENDENTE")
         r_des_r = "ENVIO PENDENTE"
    else:
    	print("Por favor, preencha sim ou não")
sleep(1)

#Gráfico
print ("Gráfico")
sleep(1)

y3_1 = float(input("Valor a"))
y3_2 = float(input("Valor b"))
y3_2 = float(input("Valor c"))
y3_4 = float(input("Valor d "))
y3_5 = float(input("Valor e"))
y3_6 = float(input("Valor f"))

#Plotagem do gráfico 
x1 = [63.5, 50.0, 38.0, 25.0, 19.0, 12.7]
y1 = [0, 10 ,40, 90, 100, 100]

x2 = [63.5, 50.0, 38.0, 25.0, 19.0, 12.7]
y2 = [0, 0 ,10, 65, 90, 95]

x3 = [63.5, 50.0, 38.0, 25.0, 19.0, 12.7]
y3 = [y3_1, y3_2, y3_2, y3_4, y3_5, y3_6]

titulo = "Composição"
eixox = "% Retido"
eixoy = "Malha"

#Legendas
plt.title(titulo)
plt.xlabel(eixox)
plt.ylabel(eixoy)

plt.plot(x1, y1, linestyle=":", color="k")
plt.plot(x2, y2, linestyle=":", color="k")
plt.plot(x3, y3, color="r")
plt.legend()
plt.savefig('{}1.png'.format(name))


workbook = xlsxwriter.Workbook('{}.xls'.format(name))

# Nomenclatura da aba
worksheet = workbook.add_worksheet('Checklist')

# Inserção do grafico na planilha
worksheet.insert_image('A19', '{}1.png'.format(name))
worksheet.insert_image('K19', '{}2.png'.format(name))

#Cabeçalho
worksheet.write(0, 0, 'nome:')
worksheet.write(0, 1, "{}".format(name))
worksheet.write(1, 0, 'CNPJ:')
worksheet.write(1, 1, "{}".format(cnpj))
worksheet.write(2, 0, 'Localidade:')
worksheet.write(2, 1, "{}".format(city))

#Linhas imutáveis
worksheet.write(4, 0, 'a')
worksheet.write(5, 0, 'b')
worksheet.write(6, 0, 'c')
worksheet.write(7, 0, 'd')
worksheet.write(8, 0, 'e')
worksheet.write(9, 0, 'f')
worksheet.write(10, 0, 'g')
worksheet.write(11, 0, 'h')
worksheet.write(12, 0, 'i')
worksheet.write(13, 0, 'j')
worksheet.write(14, 0, 'k')
worksheet.write(15, 0, 'l')
worksheet.write(16, 0, 'm')

#Colunas imutáveis
worksheet.write(3, 1, 'Resultado')
worksheet.write(3, 2, 'Parecer')

#Resultados
worksheet.write(4, 1, "{}".format(f_media))
worksheet.write(5, 1, "{}".format(f_cub))
worksheet.write(6, 1, "{}".format(m_esp))
worksheet.write(7, 1, "{}".format(ab_agua))
worksheet.write(8, 1, "{}".format(p_apar))
worksheet.write(9, 1, "{}".format(r_inter))
worksheet.write(10, 1, "{}".format(r_comp))
worksheet.write(11, 1, "{}".format(r_choque))
worksheet.write(12, 1, "{}".format(t_frag))
worksheet.write(13, 1, "{}".format(m_pulv))
worksheet.write(14, 1, "{}".format(t_arg))
worksheet.write(15, 1, "{}".format(m_uni))
worksheet.write(16, 1, "{}".format(r_des))

#Parecer 
worksheet.write(4, 2, "{}".format(f_media_r))
worksheet.write(5, 2, "{}".format(f_cub_r))
worksheet.write(6, 2, "{}".format(m_esp_r))
worksheet.write(7, 2, "{}".format(ab_agua_r))
worksheet.write(8, 2, "{}".format(p_apar_r))
worksheet.write(9, 2, "{}".format(r_inter_r))
worksheet.write(10, 2, "{}".format(r_comp_r))
worksheet.write(11, 2, "{}".format(r_choque_r))
worksheet.write(12, 2, "{}".format(t_frag_r))
worksheet.write(13, 2, "{}".format(m_pulv_r))
worksheet.write(14, 2, "{}".format(t_arg_r))
worksheet.write(15, 2, "{}".format(m_uni_r))
worksheet.write(16, 2, "{}".format(r_des_r))

workbook.close()