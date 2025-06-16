import streamlit as st
import requests
import pandas as pd
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.authentication_context import AuthenticationContext
from datetime import datetime, timedelta
import os
from zoneinfo import ZoneInfo # Importar ZoneInfo para fuso horário

# --- Configurações SharePoint ---
sharepoint_folder = '/sites/DellaVolpe/Documentos%20Compartilhados/Planejamentos/Dados_PVD/'
url_sharepoint = 'https://dellavolpecombr.sharepoint.com/sites/DellaVolpe'
username = 'marcos.silva@dellavolpe.com.br'
password = '38213824rR!!'

# --- Configurações Movidesk API ---
APITOKEN = "34779acb-809d-4628-8594-441fa68dc694"
TOP = 1000 # Quantidade máxima de tickets por requisição (para paginação)

# --- Funções SharePoint ---
def uploadSharePoint(local_file_path, sharepoint_folder):
    ctx_auth = AuthenticationContext(url_sharepoint)
    if ctx_auth.acquire_token_for_user(username, password):
        ctx = ClientContext(url_sharepoint, ctx_auth)
        with open(local_file_path, 'rb') as file_content:
            file_name = os.path.basename(local_file_path)
            target_folder = ctx.web.get_folder_by_server_relative_url(sharepoint_folder)
            target_folder.upload_file(file_name, file_content).execute_query()
            st.success(f"✅ Arquivo **{file_name}** enviado com sucesso para o SharePoint")
    else:
        st.error("❌ Autenticação no SharePoint falhou.")

# --- Funções Movidesk API (com paginação) ---
def get_tickets_page(skip, start_date_str, end_date_str):
    """
    Busca uma página de tickets da API do Movidesk.
    start_date_str e end_date_str devem ser strings formatadas (ex: "2025-01-01T00:00:00Z").
    """
    url = (
        f"https://api.movidesk.com/public/v1/tickets?"
        f"token={APITOKEN}"
        f"&$select=id,actions" # Seleciona apenas id e actions para otimização
        f"&$expand=actions"
        f"&$filter=createdDate ge {start_date_str} and createdDate le {end_date_str}"
        f"&$top={TOP}"
        f"&$skip={skip}"
    )
    try:
        response = requests.get(url)
        response.raise_for_status()  # Levanta um erro para status de erro HTTP
        data = response.json()
        if isinstance(data, dict) and "value" in data:
            return data["value"]
        elif isinstance(data, list):
            return data
        return []
    except requests.exceptions.RequestException as e:
        st.error(f"Erro na requisição à API do Movidesk: {e}")
        return []

def get_all_tickets(start_date, end_date, progress_bar=None, progress_text_placeholder=None):
    """
    Busca todos os tickets entre duas datas, usando paginação.
    start_date e end_date devem ser objetos datetime.date.
    """
    all_tickets = []
    # Usamos o mesmo filtro de time zone do Movidesk, que é UTC ('Z')
    start_dt_str = datetime.combine(start_date, datetime.min.time()).strftime("%Y-%m-%dT%H:%M:%SZ")
    end_dt_str = datetime.combine(end_date, datetime.max.time()).strftime("%Y-%m-%dT%H:%M:%SZ") # Use datetime.max.time() para o final do dia

    # Estimativa inicial de total de páginas, pode ser ajustada
    estimated_total_tickets = 1 # Começa com 1 para evitar divisão por zero
    skip = 0
    page_count = 0

    st.write("Iniciando extração de tickets...")

    while True:
        page_count += 1
        page = get_tickets_page(skip, start_dt_str, end_dt_str)
        if not page: # Se a página estiver vazia, não há mais tickets
            break

        all_tickets.extend(page)

        if progress_bar and progress_text_placeholder:
            # Atualiza a barra de progresso e o texto
            # A proporção é baseada na quantidade de dados já carregados vs. a quantidade por página
            # É uma estimativa, pois não sabemos o total exato de tickets de antemão.
            current_progress = min(1.0, (skip + len(page)) / (estimated_total_tickets + 1)) # +1 para evitar divisão por zero se for 0
            progress_bar.progress(current_progress)
            progress_text_placeholder.text(f"Extraindo página {page_count}... Total de tickets: {len(all_tickets)}")
            
            if len(page) < TOP: # Se a página retornou menos que o TOP, significa que é a última
                break
            else:
                # Se a página encheu, atualiza a estimativa de total de tickets
                estimated_total_tickets = len(all_tickets) * 2 # Assume que pode ter pelo menos mais uma página

        skip += TOP
    
    if progress_bar and progress_text_placeholder:
        progress_bar.progress(1.0) # Garante que a barra chegue a 100% no final
        progress_text_placeholder.text(f"Extração concluída! Total de tickets: {len(all_tickets)}")

    return all_tickets


# --- Nova função para extrair apenas a primeira descrição da ação ---
def extract_first_action_description(tickets):
    rows = []
    for ticket in tickets:
        ticket_id = ticket.get("id")
        actions = ticket.get("actions", [])
        
        first_action_description = None
        if actions and isinstance(actions, list) and len(actions) > 0:
            first_action_description = actions[0].get('description', None)
            
        rows.append({"TicketId": ticket_id, "FirstActionDescription": first_action_description})
    return rows

# --- Streamlit app ---
def main():
    st.title("📊 Extração Simplificada de Ações do Movidesk e Upload para SharePoint")

    # --- Seleção de data inicial ---
    data_inicial = st.date_input(
        "Selecione a data inicial para a extração:",
        value=datetime(2025, 4, 1).date(),
        min_value=datetime(2025, 1, 1).date(),
        max_value=datetime.now(ZoneInfo("America/Sao_Paulo")).date() # Data máxima é a de hoje em SP
    )

    if st.button("🚀 Iniciar extração e upload!"):
        # --- Captura o timestamp da execução ---
        execution_timestamp = datetime.now(ZoneInfo("America/Sao_Paulo")).strftime('%d/%m/%Y %H:%M:%S')
        st.info(f"🕒 Data/hora da execução: {execution_timestamp}")

        # placeholders para a barra de progresso e texto
        progress_text_placeholder = st.empty()
        progress_bar = st.progress(0)

        with st.spinner("Preparando e extraindo dados..."):
            # --- Intervalo de datas ---
            start_date = data_inicial
            end_date = datetime.now(ZoneInfo("America/Sao_Paulo")).date() # Pega a data de hoje em SP

            # Extrai todos os tickets com paginação
            all_tickets = get_all_tickets(start_date, end_date, progress_bar, progress_text_placeholder)
            
            if not all_tickets:
                st.warning("Nenhum ticket encontrado para o período selecionado ou erro na extração da API.")
                return

            # Processa os tickets para extrair a primeira descrição da ação
            actions_data = extract_first_action_description(all_tickets)
            df_final = pd.DataFrame(actions_data)

            # --- Salvando arquivo temporário ---
            csv_file_name = "TicketsMovidesk_FirstAction.csv"
            df_final.to_csv(csv_file_name, index=False)
            st.success(f"✅ Arquivo **{csv_file_name}** salvo localmente com {len(df_final)} registros.")

            # --- Upload para SharePoint ---
            uploadSharePoint(csv_file_name, sharepoint_folder)

            # --- Mostra um trecho da tabela ---
            st.subheader("Primeiras linhas dos dados extraídos:")
            st.dataframe(df_final.head())
        
        st.balloons() # Efeito de balões para sucesso

if __name__ == "__main__":
    main()
