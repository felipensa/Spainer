def juntaPDF(matricula, nome, ponto_inicial):
    import os
    from PyPDF2 import PdfFileMerger

    # GUARDA O CAMINHO E MUDA O AMBIENTE DO PYTHON PARA TAL PASTA
    pasta = fr'./Fichas {matricula} - {nome}'
    os.chdir(pasta)

    # ARMAZENA OS PDFS INDIVIDUALMENTE
    x = [a for a in os.listdir() if a.endswith(".pdf")]

    # UNIFICA OS ARQUIVOS
    merger = PdfFileMerger()
    for pdf in x:
        merger.append(open(pdf, 'rb'))

    # SALVA O ARQUIVO CONSOLIDADO
    with open("fichas_financeiras " + matricula + ".pdf", "wb") as fout:
        merger.write(fout)

    os.chdir(str(ponto_inicial))
    return 1
