import streamlit as st
import requests
import pandas as pd
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.authentication_context import AuthenticationContext
from datetime import datetime, timedelta
import os

# --- Configurações SharePoint ---
sharepoint_folder = '/sites/DellaVolpe/Documentos%20Compartilhados/Planejamentos/Dados_PVD/'
url_sharepoint = 'https://dellavolpecombr.sharepoint.com/sites/DellaVolpe'
username = 'marcos.silva@dellavolpe.com.br'
password = '38213824rR!!'

def uploadSharePoint(local_file_path, sharepoint_folder):
    ctx_auth = AuthenticationContext(url_sharepoint)
    if ctx_auth.acquire_token_for_user(username, password):
        ctx = ClientContext(url_sharepoint, ctx_auth)
        with open(local_file_path, 'rb') as file_content:
            file_name = os.path.basename(local_file_path)
            target_folder = ctx.web.get_folder_by_server_relative_url(sharepoint_folder)
            target_folder.upload_file(file_name, file_content).execute_query()
            st.success(f"✅ Arquivo **{file_name}** enviado com sucesso para o SharePoint, agora é só atualizar o indicador: https://app.powerbi.com/groups/9b59453e-21da-4a0c-8b2f-71451adc77fb/reports/5536e2b3-d355-4087-a081-2edc90854bb7/c5a41d1f74f190655f38?experience=power-bi")
    else:
        st.error("❌ Autenticação no SharePoint falhou.")

# --- Funções auxiliares ---

def get_tickets_for_date(date):
    start_of_day = date.strftime("%Y-%m-%d") + "T00:00:00.00z"
    end_of_day = date.strftime("%Y-%m-%d") + "T23:59:59.99z"
    
    api_url = (
        "https://api.movidesk.com/public/v1/tickets?"
        "token=34779acb-809d-4628-8594-441fa68dc694"
        "&$select=id,type,origin,status,urgency,type,originEmailAccount,"
        "serviceFirstLevelId,serviceFull,createdBy,owner,ownerTeam,createdDate,"
        "lastUpdate,cc,originEmailAccount,clients,actions,parentTickets,"
        "childrenTickets,statusHistories,customFieldValues,assets,chatWaitingTime,"
        "resolvedIn,subject"
        "&$expand=owner,createdBy,customFieldValues($expand=items)"
        f"&$filter=createdDate ge {start_of_day} and createdDate le {end_of_day} and ownerTeam ne 'Agente - CRC'"
    )
    
    response = requests.get(api_url)
    return response.json()

def extract_custom_fields(custom_field_values):
    custom_fields = {}
    for field in custom_field_values:
        field_id = field.get('customFieldId')
        value = field.get('value', None)

        # Verificação extra para evitar erro de 'NoneType'
        if not value and field.get('items') and isinstance(field['items'], list) and len(field['items']) > 0:
            item = field['items'][0]
            value = item.get('customFieldItem') if isinstance(item, dict) else None

        custom_fields[f'customField_{field_id}'] = value
    return custom_fields

def expand_owner(owner):
    if owner is None:
        return dict.fromkeys(['owner_id', 'owner_personType', 'owner_profileType',
                              'owner_businessName', 'owner_email', 'owner_phone', 'owner_pathPicture'], None)
    return {
        'owner_id': owner.get('id'),
        'owner_personType': owner.get('personType'),
        'owner_profileType': owner.get('profileType'),
        'owner_businessName': owner.get('businessName'),
        'owner_email': owner.get('email'),
        'owner_phone': owner.get('phone'),
        'owner_pathPicture': owner.get('pathPicture')
    }

def expand_createdby(createdby):
    if createdby is None:
        return dict.fromkeys(['createdBy_id', 'createdBy_businessName', 'createdBy_email',
                              'createdBy_phone', 'createdBy_profileType', 'createdBy_personType'], None)
    return {
        'createdBy_id': createdby.get('id'),
        'createdBy_businessName': createdby.get('businessName'),
        'createdBy_email': createdby.get('email'),
        'createdBy_phone': createdby.get('phone'),
        'createdBy_profileType': createdby.get('profileType'),
        'createdBy_personType': createdby.get('personType')
    }

def get_first_action_description(actions):
    if actions and isinstance(actions, list) and len(actions) > 0:
        return actions[0].get('description', None)
    return None

# --- Streamlit app ---

st.title("📊 Coleta de Tickets Movidesk e Upload para SharePoint")

# --- Seleção de data inicial ---
data_inicial = st.date_input(
    "Selecione a data inicial:",
    value=datetime(2025, 4, 1).date(),
    min_value=datetime(2025, 1, 1).date(),
    max_value=datetime.now().date()
)

if st.button("🚀 Iniciar a extração de dados e upload da base para atualização do indicador!"):
    # --- Captura o timestamp da execução ---
    from zoneinfo import ZoneInfo
    execution_timestamp = datetime.now(ZoneInfo("America/Sao_Paulo")).strftime('%d/%m/%Y %H:%M:%S')
    # --- Exibe o timestamp ao lado da barra de progresso ---
    st.info(f"🕒 Data/hora da execução: {execution_timestamp}")

    with st.spinner("Extraindo base..."):

        # --- Intervalo de datas ---
        start_date = datetime.combine(data_inicial, datetime.min.time())
        end_date = datetime.now()
        dates = [start_date + timedelta(days=i) for i in range((end_date - start_date).days + 1)]

        all_data = []
        progress = st.progress(0)

        for idx, date in enumerate(dates, 1):
            data = get_tickets_for_date(date)
            if isinstance(data, list):
                all_data.extend(data)
            progress.progress(idx / len(dates))

        for item in all_data:
            if 'actions' not in item:
                item['actions'] = None

        df = pd.DataFrame(all_data)
        df['first_action_description'] = df['actions'].apply(get_first_action_description)

        expanded_fields = df['customFieldValues'].apply(extract_custom_fields)
        custom_fields_df = pd.DataFrame(expanded_fields.tolist())

        expanded_owners = df['owner'].apply(expand_owner)
        owner_fields_df = pd.DataFrame(expanded_owners.tolist())

        expanded_createdBy = df['createdBy'].apply(expand_createdby)
        createdBy_fields_df = pd.DataFrame(expanded_createdBy.tolist())

        df_final = pd.concat([
            df.drop(['owner', 'customFieldValues', 'createdBy', 'actions'], axis=1), 
            owner_fields_df, 
            custom_fields_df,
            createdBy_fields_df
        ], axis=1)

        df_final['createdDate'] = pd.to_datetime(df_final['createdDate'], errors='coerce').dt.strftime('%d/%m/%Y %H:%M')
        df_final['resolvedIn'] = pd.to_datetime(df_final['resolvedIn'], errors='coerce').dt.strftime('%d/%m/%Y %H:%M')

        # --- Mapeamento dos nomes das colunas customizadas ---
        de_para_customField = {  
            'customField_177683': 'CON - ANO',  
            'customField_178151': 'CON - ID DE CUSTO',  
            'customField_177671': 'CON - MOTIVO DO TICKET',  
            'customField_177678': 'CON - PLACA',  
            'customField_178156': 'CON - PRIMEIRO CONTATO?',  
            'customField_177801': 'CON - REGIÃO',  
            'customField_177682': 'CON - TIPO DE VEICULO',  
            'customField_177804': 'CON - UF',  
            'customField_186813': 'CON - VALOR DO FRENTE NEGOCIADO?',  
            'customField_178097': 'CON - VALOR MESA',  
            'customField_177685': 'CON - VALOR NEGOCIADO',  
            'customField_174358': 'CRC - Clientes',  
            'customField_179575': 'CRC - E-mail',  
            'customField_185839': 'CRC - Filial',  
            'customField_175192': 'CRC - Motivo',  
            'customField_175190': 'CRC - Origem',  
            'customField_175170': 'CRC - Tipo',  
            'customField_175125': 'DEV - Agente Auxiliar',  
            'customField_175121': 'DEV - Classificação',  
            'customField_175118': 'DEV - Lista de Agentes',  
            'customField_177084': 'DEV - Observação',  
            'customField_175123': 'DEV - Tipo de Chamado',  
            'customField_177132': 'DEV -- Tipo de Chamado',  
            'customField_189130': 'FAT - Armazém Filial',  
            'customField_189129': 'FAT - Carregamento',  
            'customField_189128': 'FAT - CNPJ cliente sacado',  
            'customField_189139': 'FAT - Contrato cliente sacado',  
            'customField_189133': 'FAT - Grupo cliente sacado',  
            'customField_190227': 'FAT - Número do CT-e/RPS',  
            'customField_189131': 'FAT - Rota emissão',  
            'customField_189189': 'FAT - Status de finalização - Descrição',  
            'customField_189188': 'FAT - Status de finalização - Motivo',  
            'customField_189184': 'FAT - Status de finalização - Setor responsável pela divergência',  
            'customField_190231': 'FAT - Usuário',  
            'customField_188182': 'GPI - Data de Inicio',  
            'customField_188183': 'GPI - Data Final',  
            'customField_188179': 'GPI - Observação',  
            'customField_188180': 'GPI - Responsável',  
            'customField_188181': 'GPI - Status',  
            'customField_174495': 'RDA - Agente Auxiliar',  
            'customField_174494': 'RDA - Assunto',  
            'customField_174501': 'RDA - Prazo',  
            'customField_174493': 'RDA - Responsável',  
            'customField_174489': 'RDA - Tipo de Chamado',  
            'customField_187641': 'RDA - tipo de evento',  
            'customField_178941': 'RDA - Tipo de evento',  
            'customField_174488': 'RDA - Tipo de requisição',  
            'customField_174487': 'RDA - Tipo de Serviço',  
            'customField_189009': 'SAC - Categoria',  
            'customField_189005': 'SAC - Observação',  
            'customField_189008': 'SAC - Produto',  
            'customField_189007': 'SAC - Tipo de Atendimento',  
            'customField_189010': 'SAC - Tipo de Problema',  
            'customField_174486': 'SAC - Tipo de Solução',  
            'customField_174485': 'SAC - Tipo de Subproblema',  
            'customField_174484': 'SAC - Tipo de Ticket',  
            'customField_188495': 'SAC - Tipo de Ticket',  
            'customField_177674': 'SIG - Agente Auxiliar',  
            'customField_177675': 'SIG - Motivo',  
            'customField_177676': 'SIG - Observação',  
            'customField_177677': 'SIG - Responsável',  
            'customField_177673': 'SIG - Tipo de Ticket'  
        }

        df_final = df_final.rename(columns=de_para_customField)

        # --- Adiciona a coluna de timestamp ---
        df_final['execution_timestamp'] = execution_timestamp

        # --- Salvando arquivo temporário ---
        csv = 'TicketsMovidesk.csv'
        df_final.to_csv(csv, index=False)
        st.success(f"✅ Arquivo **{csv}** salvo localmente.")

        # --- Upload para SharePoint ---
        uploadSharePoint(csv, sharepoint_folder)

        # --- Mostra um trecho da tabela ---
        st.dataframe(df_final.head())

    st.balloons()
