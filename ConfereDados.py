
origem = r"C:\Users\Padrão\Desktop\BeL Tech\VJ\Royal\Planilhas"
for caminho, subpasta, arquivos in os.walk(origem):
    for contador,(nome) in enumerate(arquivos): # UMA EXECUÇÃO POR ARQUIVO
        planilha = planilhaCompleta.get_worksheet(contador)
        print(contador)
        dados = planilha.get_values()
        df = pd.DataFrame(dados).applymap(str)
        arq = caminho+"\\"+nome
        df_planilha = pd.read_excel(arq,header=None)
        df_planilha.fillna("",inplace=True)
        df_planilha.applymap(str)
        print(df)
        print(df_planilha)
        print(df.equals(df_planilha))
