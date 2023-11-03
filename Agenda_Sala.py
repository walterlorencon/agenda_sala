import io
import streamlit as st
import pandas as pd
import time
from PIL import Image
import random
import string
import datetime
import sqlite3
import streamlit.components.v1 as components
#import xlsxwriter

#Criação da Página Web-------------------------------------------------------------------------------------------------

st.set_page_config(page_title='Agendamentos',
                   page_icon="calendar:",
                   layout='wide',
                   initial_sidebar_state="collapsed"
                   )

#Conexão com o Banco de Dados------------------------------------------------------------------------------------------
con = sqlite3.connect("bdsalas.db")
cur = con.cursor()

#----------------------------------------------------------------------------------------

#Barra Lateral--------------------------------------------------------------------------------------------------------
formag = st.sidebar.form('formag', clear_on_submit = True)
vazio = formag.empty()
formag.header('REGISTRO DE AGENDAMENTOS')
tiporeg = formag.selectbox(
    'Agendamento/Cancelamento:',
    ('Agendamento', 'Cancelamento')
)

sala = formag.selectbox(
    'Sala:',
    ('Sala 01', 'Sala 02', 'Sala 03')
)

horario = formag.selectbox(
    'Horário:',
    ('07:40 - 08:30', '08:30 - 09:20', '09:20 - 10:10', '10:20 - 11:10', '11:10 - 12:00', '13:00 - 13:50', '13:50 - 14:40', '14:40 - 15:30', '15:40 - 16:30', '16:30 - 17:20', '17:20 - 18:00')

)

dtagendamento = formag.date_input('Data:')
link = formag.text_input('Link:')
senha = formag.text_input('Senha: ', type="password")
btnregistrar = formag.form_submit_button('REGISTRAR')

if btnregistrar:

    cur.execute("SELECT * FROM tbl_usuarios WHERE senha_usuario = ?", (senha,))
    dados_usuario = cur.fetchall()
      
    if len(dados_usuario) == 0:
        vazio.error('Senha Incorreta!', icon="🚨")
        time.sleep(5)
        vazio.empty()
    else:
        if tiporeg == "Agendamento":
            cod_ag = str(sala) + " / " + str(dtagendamento) + " / " + str(horario)
            cur.execute("INSERT into tbl_agenda (cod_agendamento, sala, nome_usuario, data, horario, link) values (?,?,?,?,?,?)", (cod_ag, sala, dados_usuario[0][1], dtagendamento, horario, link))
            con.commit()
            vazio.success('Agendamento Realizado com Sucesso!', icon="✅")
            time.sleep(5)
            vazio.empty()
        
        else:
            cod_ag = str(sala) + " / " + str(dtagendamento) + " / " + str(horario)
            cur.execute("SELECT nome_usuario FROM tbl_agenda WHERE cod_agendamento = ?", (cod_ag,))
            usuario_agendamento = cur.fetchall()
            if (usuario_agendamento[0][0] == dados_usuario[0][1]) or (dados_usuario[0][3] == "Administrador"):
                cur.execute("DELETE FROM tbl_agenda WHERE cod_agendamento = ?", (cod_ag,))
                con.commit()
                vazio.success('Agendamento Cancelado com Sucesso!', icon="✅")
                time.sleep(5)
                vazio.empty()
            else:
                vazio.error('Cancelamento Não Autorizado', icon="🚨")
                time.sleep(5)
                vazio.empty()

### Página Principal--------------------------------------------------------------------------------------------------

st.title('📅 RESERVA DE SALAS')

tab1, tab2, tab3 = st.tabs(['*🏠 INÍCIO','🔍 CONSULTA AGENDA', '👤 CADASTRO DE USUÁRIOS*'])


# Aba Início-----------------------------------------------------------------------------------------------------

with tab1:
    st.header('Verificar Agendamentos')
    st.markdown(
        """
        O presente painel foi desenvolvido para otimizar e dar eficiência nos agendamentos das salas de informática, 
         integrando tecnologia e informação para o desenvolvimento dos alunos.
         
        Utilize a barra lateral para agendar uma sala. Clique no sinal de > no canto superior esquerdo da tela para abrir a barra lateral.
        
    """
    )
    #imagem = Image.open(r'C:\Users\55189\Desktop\Teste\La e Do.jpg')
    #st.image(imagem)

# Aba Consulta Agenda--------------------------------------------------------------------------------------------------
with tab2:
    hoje = datetime.date.today()
    dados = pd.read_sql("SELECT * FROM tbl_agenda", con).set_index(['data'])
    agenda = dados.drop(dados.columns[0], axis=1)
    agenda.sort_values(by=['data'], inplace=True, ascending=False)
    agenda.rename(columns={"sala":"SALA RESERVADA"}, inplace=True)
    agenda.rename(columns={"nome_usuario": "RESPONSÁVEL PELA RESERVA"}, inplace=True)
    agenda.index.name = 'DATA DA RESERVA'
    agenda.rename(columns={"horario": "HORÁRIO RESERVADO"}, inplace=True)
    

    
    # download button 1 to download dataframe as csv
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
    # Write each dataframe to a different worksheet.
        agenda.to_excel(writer, sheet_name='Sheet1', index=False)
        writer.save()
        download1 = st.download_button(
            label="Exportar para Excel",
            data=buffer,
            file_name='Relatório.xlsx',
            mime='application/vnd.ms-excel'
        )

    st.data_editor(
    agenda,
    column_config={
        "link": st.column_config.LinkColumn(
            "ACESSO",
            help="Link para acessar a reunião",
        )
    },
    hide_index=True,
    )   
    #tabela = st.write(agenda, unsafe_allow_html=True )
# Codigo Java Script-------------------------------------------------------------------------------------------------

 # Aba Cadastro de Usuários---------------------------------------------------------------------------------------------
with tab3:
    cad_usuario = tab3.form('cad_usuario', clear_on_submit=True)
    msgs = cad_usuario.empty()
    cpf = cad_usuario.text_input('Cadastro de CPF (somente números): ')
    nome = cad_usuario.text_input('Cadastro de Nome: ')
    cadsenha = cad_usuario.text_input('Cadastro de Senha: ')
    perfil = cad_usuario.selectbox('Cadastro de Perfil: ',('Usuário', 'Administrador'))
    senhaadm = cad_usuario.text_input('Senha de Administrador: ', type="password")
    btncaduser = cad_usuario.form_submit_button('CADASTRAR USUÁRIO')
    btnapagauser = cad_usuario.form_submit_button('EXCLUIR USUÁRIO')

    if btncaduser:
        cur.execute("SELECT * FROM tbl_usuarios WHERE senha_usuario = ?", (senhaadm,))
        dados_usuario = cur.fetchall()

        if dados_usuario[0][3] == "Administrador":
            cur.execute("INSERT into tbl_usuarios (cpf, nome_usuario, senha_usuario, perfil_usuario) values (?,?,?,?)", (cpf, nome, cadsenha, perfil,))
            con.commit()
        else:
            msgs.error('Sem Privilégios para Cadastrar Novos Usuários', icon="🚨")
            time.sleep(5)
            msgs.empty()

    if btnapagauser:
        cur.execute("SELECT * FROM tbl_usuarios WHERE senha_usuario = ?", (senhaadm,))
        dados_usuario = cur.fetchall()

        if dados_usuario[0][3] == "Administrador":
            cur.execute("DELETE FROM tbl_usuarios WHERE cpf = ?", (cpf,))
            con.commit()
            msgs.success('Usuário Excluído com Sucesso!', icon="✅")
            time.sleep(5)
            vazio.empty()
        else:
            msgs.error('Sem Privilégios para Excluir Usuários', icon="🚨")
            time.sleep(5)
            msgs.empty()

#Fim-----------------------------------------------------------------
