import openpyxl
from PIL import Image, ImageDraw, ImageFont

workbook_alunos = openpyxl.load_workbook('planilha.xlsx')
sheet_alunos = workbook_alunos['Planilha1']

for indice, line in enumerate(sheet_alunos.iter_rows(min_row=2)):
    nome_aluno = line[0].value
    nome_curso = line[1].value
    tipo_participacao = line[2].value
    data_inicio = line[3].value
    data_fim = line[4].value
    carga_horaria = line[5].value
    data_emissao = line[6].value

    font_name = ImageFont.truetype('./Morganite-Bold.ttf', size=60)
    font_geral = ImageFont.truetype('./Morganite-Light.ttf', size=52)
    font_datas = ImageFont.truetype('./Morganite-Light.ttf', size=34)

    image = Image.open('./certificado.png')
    desenhar = ImageDraw.Draw(image)

    desenhar.text((376, 288), nome_aluno, fill='black', font=font_name)
    desenhar.text((390, 336), nome_curso, fill='black', font=font_geral)
    desenhar.text((520, 378), tipo_participacao, fill='black', font=font_geral)
    desenhar.text((539, 421), str(carga_horaria)+" Horas", fill='black', font=font_geral)
    data_inicio_str = data_inicio.strftime("%d-%m-%Y %H:%M:%S")
    desenhar.text((285, 644), str(data_inicio_str[:10]), fill='blue', font=font_datas)
    data_fim_str = data_fim.strftime("%d-%m-%Y %H:%M:%S")
    desenhar.text((285, 699), str(data_fim_str[:10]), fill='blue', font=font_datas)
    data_emissao_str = data_emissao.strftime("%d-%m-%Y %H:%M:%S")
    desenhar.text((820, 698), str(data_emissao_str[:10]), fill='blue', font=font_datas)

    image.save(f'./testes/{indice} {nome_aluno} certificados.png')