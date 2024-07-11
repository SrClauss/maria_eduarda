import pandas
import tinydb
import multiprocessing




print("Carregando DFS")
dfs = [pandas.read_excel('data.xlsx', sheet_name=i.__str__()) for i in range(1, 7)]
print("DFS carregados")

db = tinydb.TinyDB('db.json')
table = db.table('hemogramas')
table_erros = db.table('erros')
for df_num in dfs:
    print("iniciando df")
    cod_atendimentos = list(set(df_num['cod_atendimento'].tolist()))
    cod_atendimentos.sort()


    cod1 = df_num.iloc[0]['cod_atendimento']



    for cod in cod_atendimentos:
        print(cod)
        df = df_num[df_num['cod_atendimento'] == cod]

        if len(df) != 7:    
            print('erro', cod)
            table_erros.insert({'cod_atendimento': cod, 'erro': 'erro de tamanho'})
        
            
        obj = {}
        obj['cod_atendimento'] = cod
        obj['origem'] = df.iloc[0]['origem']
        obj['men'] = df.iloc[0]['men']
        obj['exame'] = df.iloc[0]['exame']
        obj['idade'] = df.iloc[0]['idade'].split(' ')[0]
        obj['sexo'] = df.iloc[0]['sexo']
        obj['data_exame'] = df.iloc[0]['data_exame'].__str__()
        try:
            obj['hemacias'] = df.iloc[0]['resultado']
        except:
            obj['hemacias'] = ''
        try:
            obj['hemoglobina'] = df.iloc[1]['resultado']
        except:
            obj['hemoglobina'] = ''     

        try:
            obj['CHCM'] = df.iloc[2]['resultado']
        except:
            obj['CHCM'] = ''

        try:
            obj['hematocrito'] = df.iloc[3]['resultado']
        except:
            obj['hematocrito'] = ''
        try:
            obj['VCM'] = df.iloc[4]['resultado']
        except:
            obj['VCM'] = ''
        try:
            obj['HCM'] = df.iloc[5]['resultado']
        except:
            obj['HCM'] = ''
        try:   
            obj['RDW'] = df.iloc[6]['resultado']
        except:
            obj['RDW'] = ''

        



        table.insert(obj)







