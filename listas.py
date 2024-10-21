import streamlit as st
from datetime import time
import json
import excel

def save_data(form_data, filename):
    try:
        with open(filename, 'w') as f:
            json.dump(form_data, f, indent=4)
        st.success(f"Dados salvos com sucesso no arquivo. Pode realizar o Download da planilha preenchida.")
    except Exception as e:
        st.error(f"Erro ao salvar os dados: {e}")

# Função para lista 1
def lista1_form():
    st.title("Preenchimento: Lista de Capacitação Online - ÁREA MEIO (GEP)")
    
    # Utilizando st.form para agrupar o formulário e evitar reset ao apertar Enter
    with st.form(key="form_lista1"):
        # Campos temporários para serem submetidos com on_change vazio para evitar submissão ao pressionar Enter
        tema_input = st.text_input("Tema")
        palestrante_input = st.text_input("Nome do(a) palestrante")
        publico_alvo_input = st.text_input("Público Alvo")
        data_input = st.date_input("Data")
        horario_inicio_input = st.text_input("Horário de Início")
        horario_fim_input = st.text_input("Horário de Fim")
        carga_horaria_input = st.text_input("Carga Horária")
        plataforma_input = st.text_input("Plataforma Online")
        ct_gestao_input = st.text_input("Contrato de Gestão")
        # Criando duas colunas para organizar os botões
        col1, col2 = st.columns([2, 1])  # Coluna 1 maior que a coluna 2

        with col1:
            # Botão de envio dentro do formulário
            submit_button = st.form_submit_button(label="Enviar")

        with col2:
            # Botão para fazer o download
            download_button = st.form_submit_button(label="Faça o download aqui")
    if submit_button:
        form_data = {
            "Tema": tema_input,
            "Palestrante": palestrante_input,
            "Publico Alvo": publico_alvo_input,
            "Data": str(data_input),
            "Horario de Inicio": horario_inicio_input,
            "Horario de Fim": horario_fim_input,
            "Carga Horaria": carga_horaria_input,
            "Plataforma Online": plataforma_input,
            "Contrato de Gestao": ct_gestao_input
        }
        save_data(form_data, 'lista1_data.json')
        st.session_state['form_enviado'] = True

    if download_button:
        if st.session_state.get('form_enviado', False):
            excel.preencher_excel_com_json('lista1_data.json', 'Lista_Preenchida.xlsx')
        else:
            st.error("Por favor, preencha as informações e clique no botão de enviar antes de fazer o download.")

# Função para lista 2
def lista2_form():
    st.title("Preenchimento: Lista de Capacitação Presencial - ÁREA MEIO (GEP)")
    # Utilizando st.form para agrupar o formulário e evitar reset ao apertar Enter
    with st.form(key="form_lista2"):
        # Campos temporários para serem submetidos com on_change vazio para evitar submissão ao pressionar Enter
        tema_input = st.text_input("Tema")
        palestrante_input = st.text_input("Nome do(a) palestrante")
        publico_alvo_input = st.text_input("Público Alvo")
        data_input = st.date_input("Data")
        horario_inicio_input = st.text_input("Horário de Início")
        horario_fim_input = st.text_input("Horário de Fim")
        carga_horaria_input = st.text_input("Carga Horária")
        local_input = st.text_input("Local")
        ct_gestao_input = st.text_input("Contrato de Gestão")
            # Criando duas colunas para organizar os botões
        col1, col2 = st.columns([2, 1])  # Coluna 1 maior que a coluna 2

        with col1:
            # Botão de envio dentro do formulário
            submit_button = st.form_submit_button(label="Enviar")

        with col2:
            # Botão para fazer o download
            download_button = st.form_submit_button(label="Faça o download aqui")
    if submit_button:
        form_data = {
            "Tema": tema_input,
            "Palestrante": palestrante_input,
            "Publico Alvo": publico_alvo_input,
            "Data": str(data_input),
            "Horario de Inicio": horario_inicio_input,
            "Horario de Fim": horario_fim_input,
            "Carga Horaria": carga_horaria_input,
            "Local": local_input,
            "Contrato de Gestao": ct_gestao_input
        }
        save_data(form_data, 'lista2_data.json')
        st.session_state['form_enviado'] = True
    if download_button:
        if st.session_state.get('form_enviado', False):
            excel.preencher_excel_com_json('lista2_data.json', 'Lista_Preenchida.xlsx')
        else:
            st.error("Por favor, preencha as informações e clique no botão de enviar antes de fazer o download.")


# Função para lista 3
def lista3_form():
    st.title("Preenchimento: Lista de Capacitação Presencial - EMESP (GEP)")
    # Utilizando st.form para agrupar o formulário e evitar reset ao apertar Enter
    with st.form(key="form_lista3"):
        # Campos temporários para serem submetidos com on_change vazio para evitar submissão ao pressionar Enter
        tema_input = st.text_input("Tema")
        palestrante_input = st.text_input("Nome do(a) palestrante")
        publico_alvo_input = st.text_input("Público Alvo")
        data_input = st.date_input("Data")
        horario_inicio_input = st.text_input("Horário de Início")
        horario_fim_input = st.text_input("Horário de Fim")
        carga_horaria_input = st.text_input("Carga Horária")
        local_input = st.text_input("Local")
        ct_gestao_input = st.text_input("Contrato de Gestão")
        # Criando duas colunas para organizar os botões
        col1, col2 = st.columns([2, 1])  # Coluna 1 maior que a coluna 2

        with col1:
            # Botão de envio dentro do formulário
            submit_button = st.form_submit_button(label="Enviar")

        with col2:
            # Botão para fazer o download
            download_button = st.form_submit_button(label="Faça o download aqui")
    if submit_button:
        form_data = {
            "Tema": tema_input,
            "Palestrante": palestrante_input,
            "Publico Alvo": publico_alvo_input,
            "Data": str(data_input),
            "Horario de Inicio": horario_inicio_input,
            "Horario de Fim": horario_fim_input,
            "Carga Horaria": carga_horaria_input,
            "Local": local_input,
            "Contrato de Gestao": ct_gestao_input
        }
        save_data(form_data, 'lista3_data.json')
        st.session_state['form_enviado'] = True
    if download_button:
        if st.session_state.get('form_enviado', False):
            excel.preencher_excel_com_json('lista3_data.json', 'Lista_Preenchida.xlsx')
        else:
            st.error("Por favor, preencha as informações e clique no botão de enviar antes de fazer o download.")


# Função para lista 4
def lista4_form():
    st.title("Preenchimento: Lista de Capacitação Online - EMESP (GEP)")
    # Utilizando st.form para agrupar o formulário e evitar reset ao apertar Enter
    with st.form(key="form_lista4"):
        # Campos temporários para serem submetidos com on_change vazio para evitar submissão ao pressionar Enter
        tema_input = st.text_input("Tema")
        palestrante_input = st.text_input("Nome do(a) palestrante")
        publico_alvo_input = st.text_input("Público Alvo")
        data_input = st.date_input("Data")
        horario_inicio_input = st.text_input("Horário de Início")
        horario_fim_input = st.text_input("Horário de Fim")
        carga_horaria_input = st.text_input("Carga Horária")
        plataforma_input = st.text_input("Plataforma Online")
        ct_gestao_input = st.text_input("Contrato de Gestão")

        # Criando duas colunas para organizar os botões
        col1, col2 = st.columns([2, 1])  # Coluna 1 maior que a coluna 2

        with col1:
            # Botão de envio dentro do formulário
            submit_button = st.form_submit_button(label="Enviar")

        with col2:
            # Botão para fazer o download
            download_button = st.form_submit_button(label="Faça o download aqui")
    if submit_button:
        form_data = {
            "Tema": tema_input,
            "Palestrante": palestrante_input,
            "Publico Alvo": publico_alvo_input,
            "Data": str(data_input),
            "Horario de Inicio": horario_inicio_input,
            "Horario de Fim": horario_fim_input,
            "Carga Horaria": carga_horaria_input,
            "Plataforma Online": plataforma_input,
            "Contrato de Gestao": ct_gestao_input
        }
        save_data(form_data, 'lista4_data.json')
        st.session_state['form_enviado'] = True
    if download_button:
        if st.session_state.get('form_enviado', False):
            excel.preencher_excel_com_json('lista4_data.json', 'Lista_Preenchida.xlsx')
        else:
            st.error("Por favor, preencha as informações e clique no botão de enviar antes de fazer o download.")
# Função para lista 5
def lista5_form():
    st.title("Preenchimento: Lista de Capacitação Online - Guri (GEP)")
    # Utilizando st.form para agrupar o formulário e evitar reset ao apertar Enter
    with st.form(key="form_lista5"):
        # Campos temporários para serem submetidos com on_change vazio para evitar submissão ao pressionar Enter
        tema_input = st.text_input("Tema")
        palestrante_input = st.text_input("Nome do(a) palestrante")
        publico_alvo_input = st.text_input("Público Alvo")
        data_input = st.date_input("Data")
        horario_inicio_input = st.text_input("Horário de Início")
        horario_fim_input = st.text_input("Horário de Fim")
        carga_horaria_input = st.text_input("Carga Horária")
        plataforma_input = st.text_input("Plataforma Online")
        ct_gestao_input = st.text_input("Contrato de Gestão")
        
        # Criando duas colunas para organizar os botões
        col1, col2 = st.columns([2, 1])  # Coluna 1 maior que a coluna 2

        with col1:
            # Botão de envio dentro do formulário
            submit_button = st.form_submit_button(label="Enviar")

        with col2:
            # Botão para fazer o download
            download_button = st.form_submit_button(label="Faça o download aqui")
    if submit_button:
        form_data = {
            "Tema": tema_input,
            "Palestrante": palestrante_input,
            "Publico Alvo": publico_alvo_input,
            "Data": str(data_input),
            "Horario de Inicio": horario_inicio_input,
            "Horario de Fim": horario_fim_input,
            "Carga Horaria": carga_horaria_input,
            "Plataforma Online": plataforma_input,
            "Contrato de Gestao": ct_gestao_input
        }
        save_data(form_data, 'lista5_data.json')
        st.session_state['form_enviado'] = True
    if download_button:
        if st.session_state.get('form_enviado', False):
            excel.preencher_excel_com_json('lista5_data.json', 'Lista_Preenchida.xlsx')
        else:
            st.error("Por favor, preencha as informações e clique no botão de enviar antes de fazer o download.")


# Função para lista 6
def lista6_form():
    st.title("Preenchimento: Lista de Capacitação Presencial - Guri (GEP)")
    # Utilizando st.form para agrupar o formulário e evitar reset ao apertar Enter
    with st.form(key="form_lista6"):
        # Campos temporários para serem submetidos com on_change vazio para evitar submissão ao pressionar Enter
        tema_input = st.text_input("Tema")
        palestrante_input = st.text_input("Nome do(a) palestrante")
        publico_alvo_input = st.text_input("Público Alvo")
        data_input = st.date_input("Data")
        horario_inicio_input = st.text_input("Horário de Início")
        horario_fim_input = st.text_input("Horário de Fim")
        carga_horaria_input = st.text_input("Carga Horária")
        local_input = st.text_input("Local")
        ct_gestao_input = st.text_input("Contrato de Gestão")

        # Criando duas colunas para organizar os botões
        col1, col2 = st.columns([2, 1])  # Coluna 1 maior que a coluna 2

        with col1:
            # Botão de envio dentro do formulário
            submit_button = st.form_submit_button(label="Enviar")

        with col2:
            # Botão para fazer o download
            download_button = st.form_submit_button(label="Faça o download aqui")

    if submit_button:
        form_data = {
            "Tema": tema_input,
            "Palestrante": palestrante_input,
            "Publico Alvo": publico_alvo_input,
            "Data": str(data_input),
            "Horario de Inicio": horario_inicio_input,
            "Horario de Fim": horario_fim_input,
            "Carga Horaria": carga_horaria_input,
            "Local": local_input,
            "Contrato de Gestao": ct_gestao_input
        }
        save_data(form_data, 'lista6_data.json')
        st.session_state['form_enviado'] = True
    if download_button:
        if st.session_state.get('form_enviado', False):
            excel.preencher_excel_com_json('lista6_data.json', 'Lista_Preenchida.xlsx')
        else:
            st.error("Por favor, preencha as informações e clique no botão de enviar antes de fazer o download.")




# Função para lista 7
def lista7_form():
    st.title("Preenchimento: Lista de Atividades Extraclasse e Nova Profissões (PED. GURI)")
    # Utilizando st.form para agrupar o formulário e evitar reset ao apertar Enter
    with st.form(key="form_lista7"):
        # Campos temporários para serem submetidos com on_change vazio para evitar submissão ao pressionar Enter
        temas = ['Masterclass', 'Workshop', 'Studio Classe', 'Formação de Profissionais da Cultura, Nova Profssiões Musicais e Empreendedorismo', 'Festival Multicultural']
        tema_input = st.selectbox('Qual o tema da atividade?', temas)

        programas = ['Guri Capital e Grande São Paulo', 'Guri Interior, Fundação CASA e Litoral']
        programa_input = st.selectbox('Qual o Programa relacionado à ativiadde?', programas)

        grupo_input = st.text_input('Qual é o grupo participante da atividade?')

        atividade_input = st.text_input("Título da Atividade")

        polos_input = st.text_input("Polos Participantes")

        data_input = st.date_input("Data")

        horario_inicio_input = st.text_input("Horário de Início")

        horario_fim_input = st.text_input("Horário de Fim")

        local_input = st.text_input("Local")

        professor_input = st.text_input("Professor(a)/Educador(a)/Convidado(a)")

        # Criando duas colunas para organizar os botões
        col1, col2 = st.columns([2, 1])  # Coluna 1 maior que a coluna 2

        with col1:
            # Botão de envio dentro do formulário
            submit_button = st.form_submit_button(label="Enviar")

        with col2:
            # Botão para fazer o download
            download_button = st.form_submit_button(label="Faça o download aqui")
    if submit_button:
        form_data = {
            "Tema": tema_input,
            "Programa": programa_input,
            "Nome da Atividade": atividade_input,
            "Grupo": grupo_input,
            "Data": str(data_input),
            "Horario de Inicio": horario_inicio_input,
            "Horario de Fim": horario_fim_input,
            "Polos Participantes": polos_input,
            "Local": local_input,
            "Professor": professor_input
        }
        save_data(form_data, 'lista7_data.json')
        st.session_state['form_enviado'] = True
    if download_button:
        if st.session_state.get('form_enviado', False):
            excel.preencher_excel_com_json('lista7_data.json', 'Lista_Preenchida.xlsx')
        else:
            st.error("Por favor, preencha as informações e clique no botão de enviar antes de fazer o download.")

# Função para lista 8
def lista8_form():
    st.title("Preenchimento: Lista de Grupos de Polo (PED. GURI)")
    # Utilizando st.form para agrupar o formulário e evitar reset ao apertar Enter
    with st.form(key="form_lista8"):
        # Campos temporários para serem submetidos com on_change vazio para evitar submissão ao pressionar Enter
        temas = ['Grupos Musicais dos Polos']
        tema_input = st.selectbox('Qual o tema da Atividade?', temas)

        programas = ['Guri Capital e Grande São Paulo', 'Guri Interior, Fundação CASA e Litoral']
        programa_input = st.selectbox('Qual o Programa relacionado à ativiadde?', programas)

        atividade_input = st.text_input("Título da Atividade")

        polos_input = st.text_input("Polos Participantes")

        data_input = st.date_input("Data")

        horario_inicio_input = st.text_input("Horário de Início")

        horario_fim_input = st.text_input("Horário de Fim")

        local_input = st.text_input("Local")

        professor_input = st.text_input("Professor(a)/Educador(a)/Convidado(a)")

        # Criando duas colunas para organizar os botões
        col1, col2 = st.columns([2, 1])  # Coluna 1 maior que a coluna 2

        with col1:
            # Botão de envio dentro do formulário
            submit_button = st.form_submit_button(label="Enviar")

        with col2:
            # Botão para fazer o download
            download_button = st.form_submit_button(label="Faça o download aqui")
    if submit_button:
        form_data = {
            "Tema": tema_input,
            "Programa": programa_input,
            "Nome da Atividade": atividade_input,
            "Data": str(data_input),
            "Horario de Inicio": horario_inicio_input,
            "Horario de Fim": horario_fim_input,
            "Polos Participantes": polos_input,
            "Local": local_input,
            "Professor": professor_input
        }
        save_data(form_data, 'lista8_data.json')
        st.session_state['form_enviado'] = True
    if download_button:
        if st.session_state.get('form_enviado', False):
            excel.preencher_excel_com_json('lista8_data.json', 'Lista_Preenchida.xlsx')
        else:
            st.error("Por favor, preencha as informações e clique no botão de enviar antes de fazer o download.")


# Função para lista 9
def lista9_form():
    st.title("Preenchimento: Lista de Presença - GURI (SOCIAL)")
    # Utilizando st.form para agrupar o formulário e evitar reset ao apertar Enter
    with st.form(key="form_lista9"):
        # Campos temporários para serem submetidos com on_change vazio para evitar submissão ao pressionar Enter
        metas = ['Atividades Culturais',
                 'Encontro de Rede Socioterritorial',
                 'Oficina Socioeducativa com as Famílias da Crianças e Adolescentes',
                 'Oficina Socioeducativa com Crianças e Adolescentes',
                 'Oficina Socioeducativa de Integração com os Polos',
                 'Projeto com Famílias - Economia Solidária',
                 'Projeto de Vida - Trilhas e Carreiras',
                 'Projeto Guri Participativo - Protagonismo e Participação',
                 'Projeto Socializando',
                 'Projeto Temático Relacionado aos Objetivos de Desenvolvimento Sustentável',
                 'Reunião com Famílias'
                 ]
        meta_input = st.selectbox('Qual a meta relacionada à atividade?', metas)

        tema_input = st.text_input('Qual o tema da atividade?')

        publico_alvo_input = st.text_input("Qual o público alvo da atividade?")

        convidado_input = st.text_input("Qual o nome do Convidado?")

        polos_input = st.text_input("Qual(is) o(s) polo(s) participante(s) da atividade?")

        data_input = st.date_input("Data")

        horario_inicio_input = st.text_input("Horário de Início")

        horario_fim_input = st.text_input("Horário de Fim")

        local_input = st.text_input("Local")

        responsavel_input = st.text_input("Responsável pela oficina ou atividade")

        regional_input = st.text_input('Regional')

        tipo = ['PRESENCIALMENTE', 'REMOTAMENTE']
        tipo_input = st.selectbox('A ação foi realizada de que maneira?', tipo)

        carga_horaria_input = st.text_input('Carga Horária da Atividade')

        # Criando duas colunas para organizar os botões
        col1, col2 = st.columns([2, 1])  # Coluna 1 maior que a coluna 2

        with col1:
            # Botão de envio dentro do formulário
            submit_button = st.form_submit_button(label="Enviar")

        with col2:
            # Botão para fazer o download
            download_button = st.form_submit_button(label="Faça o download aqui")
    if submit_button:
        form_data = {
            "Meta": meta_input,
            "Tema": tema_input,
            "Publico Alvo": publico_alvo_input,
            "Convidado": convidado_input,
            "Polo": polos_input,
            "Data": str(data_input),
            "Horario de Inicio": horario_inicio_input,
            "Horario de Fim": horario_fim_input,
            "Responsável pela Atividade": responsavel_input,
            "Local": local_input,
            "Regional": regional_input,
            "Remoto ou Presencial": tipo_input,
            "Carga Horaria": carga_horaria_input
        }
        save_data(form_data, 'lista9_data.json')
        st.session_state['form_enviado'] = True
    if download_button:
        if st.session_state.get('form_enviado', False):
            excel.preencher_excel_com_json('lista9_data.json', 'Lista_Preenchida.xlsx')
        else:
            st.error("Por favor, preencha as informações e clique no botão de enviar antes de fazer o download.")


# Função para lista 10
def lista10_form():
    st.title("Preenchimento: Lista de Presença - EMESP (SOCIAL)")
    # Utilizando st.form para agrupar o formulário e evitar reset ao apertar Enter
    with st.form(key="form_lista10"):
        # Campos temporários para serem submetidos com on_change vazio para evitar submissão ao pressionar Enter
        metas = ['Ação Socioeducativa sobre os Objetivos de Desenvolvimento Sustentável', 
                 'Atividades Culturais',
                 'Encontro da Rede - Articulaçã no Território',
                 'Oficina Socioeducativa com Aluno(as)',
                 'Oficina Socioeducativa com Família(as)',
                 'Oficina Socioeducativa para os(as) Grupo Artísticos e Academia de Ópera',
                 'Projeto Socializando',
                 'Reunião com Famílias'
                 ]
        meta_input = st.selectbox('Qual a meta relacionada à atividade?', metas)
        tema_input = st.text_input('Qual o tema da atividade?')
        publico_alvo_input = st.text_input("Qual o público alvo da atividade?")
        convidado_input = st.text_input("Qual o nome do Convidado?")
        data_input = st.date_input("Data")
        horario_inicio_input = st.text_input("Horário de Início")
        horario_fim_input = st.text_input("Horário de Fim")
        local_input = st.text_input("Local")
        responsavel_input = st.text_input("Responsável pela oficina ou atividade")
        parceiros_input = st.text_input('Parceiros')

        tipo = ['PRESENCIALMENTE', 'REMOTAMENTE']
        tipo_input = st.selectbox('A ação foi realizada de que maneira?', tipo)
        carga_horaria_input = st.text_input('Carga Horária da Atividade')

        # Criando duas colunas para organizar os botões
        col1, col2 = st.columns([2, 1])  # Coluna 1 maior que a coluna 2

        with col1:
            # Botão de envio dentro do formulário
            submit_button = st.form_submit_button(label="Enviar")

        with col2:
            # Botão para fazer o download
            download_button = st.form_submit_button(label="Faça o download aqui")
    if submit_button:
        form_data = {
            "Meta": meta_input,
            "Tema": tema_input,
            "Público Alvo": publico_alvo_input,
            "Convidado": convidado_input,
            "Data": str(data_input),
            "Horario de Inicio": horario_inicio_input,
            "Horario de Fim": horario_fim_input,
            "Responsavel pela Atividade": responsavel_input,
            "Parceiros": parceiros_input,
            "Local": local_input,
            "Remoto ou Presencial": tipo_input,
            "Carga Horaria": carga_horaria_input
        }
        save_data(form_data, '10lista.json')
        st.session_state['form_enviado'] = True
    if download_button:
        if st.session_state.get('form_enviado', False):
            excel.preencher_excel_com_json('10lista.json', 'Lista_Preenchida.xlsx')
        else:
            st.error("Por favor, preencha as informações e clique no botão de enviar antes de fazer o download.")