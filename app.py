import streamlit as st
from listas import lista1_form, lista2_form, lista3_form, lista4_form, lista5_form, lista6_form, lista7_form, lista8_form, lista9_form, lista10_form

# Centraliza o título da aplicação
st.markdown("<h1 style='text-align: center;'>Ambiente de Preenchimento das Listas de Presença</h1>", unsafe_allow_html=True)
st.markdown("<h5 style='text-align: justify;'>Para preencher a lista de presença, selecione sua área específica na barra lateral à esquerda. Em seguida, escolha a lista desejada no filtro abaixo e clique no botão para prosseguir com o preenchimento</h5>", unsafe_allow_html=True)

# Dicionário com as listas
dicionario = {
    'GESTÃO DE PESSOAS': [
        'Lista de Capacitação Online - ÁREA MEIO (GEP)',
        'Lista de Capacitação Presencial - ÁREA MEIO (GEP)',
        'Lista de Capacitação Online - EMESP (GEP)',
        'Lista de Capacitação Presencial - EMESP (GEP)',
        'Lista de Capacitação Online - Guri (GEP)',
        'Lista de Capacitação Presencial - Guri (GEP)'
    ],
    'PEDAGÓGICO GURI': [
        'Lista de Atividades Extraclasse e Nova Profissões (PED. GURI)',
        'Lista de Grupos de Polo (PED. GURI)'
    ],
    'SOCIAL': [
        'Lista de Presença - GURI (SOCIAL)',
        'Lista de Presença - EMESP (SOCIAL)'
    ]
}

# Filtro de áreas com selectbox na sidebar (seleção única)
st.sidebar.subheader("Filtro de Área: Escolha a Área de Interesse")

# Seleção de uma única área
selected_area = st.sidebar.selectbox("Selecione uma área", list(dicionario.keys()))

# Seleção da lista com base na área selecionada
listas_disponiveis = dicionario[selected_area]

# Adicionando uma opção para selecionar uma lista
selected_lista = st.sidebar.radio("Escolha uma lista:", listas_disponiveis)

# Navegação para a página correspondente
if selected_lista == 'Lista de Capacitação Online - ÁREA MEIO (GEP)':
    lista1_form()  
elif selected_lista == 'Lista de Capacitação Presencial - ÁREA MEIO (GEP)':
    lista2_form()  
elif selected_lista == 'Lista de Capacitação Presencial - EMESP (GEP)':
    lista3_form()  
elif selected_lista == 'Lista de Capacitação Online - EMESP (GEP)':
    lista4_form()  
elif selected_lista == 'Lista de Capacitação Online - Guri (GEP)':
    lista5_form()  
elif selected_lista == 'Lista de Capacitação Presencial - Guri (GEP)':
    lista6_form()  
elif selected_lista == 'Lista de Atividades Extraclasse e Nova Profissões (PED. GURI)':
    lista7_form()  
elif selected_lista == 'Lista de Grupos de Polo (PED. GURI)':
    lista8_form()  
elif selected_lista == 'Lista de Presença - GURI (SOCIAL)':
    lista9_form()  
elif selected_lista == 'Lista de Presença - EMESP (SOCIAL)':
    lista10_form()
