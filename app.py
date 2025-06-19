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

APITOKEN = "34779acb-809d-4628-8594-441fa68dc694"
TOP = 1000

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

def get_tickets_page(skip, start_date, end_date):
    url = (
        f"https://api.movidesk.com/public/v1/tickets?"
        f"token={APITOKEN}"
        f"&$select=id,actions"
        f"&$expand=actions"
        f"&$filter=createdDate ge {start_date} and createdDate le {end_date}"
        f"&$top={TOP}"
        f"&$skip={skip}"
    )
    response = requests.get(url)
    response.raise_for_status()
    data = response.json()
    if isinstance(data, dict) and "value" in data:
        return data["value"]
    elif isinstance(data, list):
        return data
    return []

def get_all_tickets(start_date, end_date):
    skip = 0
    all_tickets = []
    while True:
        page = get_tickets_page(skip, start_date, end_date)
        all_tickets.extend(page)
        if len(page) < TOP:
            break
        skip += TOP
    return all_tickets

def extract_actions(tickets):
    rows = []
    for ticket in tickets:
        ticket_id = ticket.get("id")
        actions = ticket.get("actions", [])
        if not actions:
            rows.append({"TicketId": ticket_id, **{}})
        else:
            for action in actions:
                row = {"TicketId": ticket_id}
                for k, v in action.items():
                    row[f"Action_{k}"] = v
                rows.append(row)
    return rows

def main():
    st.title("📊 Extração de Tickets Movidesk com paginação e upload para SharePoint")

    data_inicial = st.date_input(
        "Data inicial:",
        value=datetime(2025,4,1).date(),
        min_value=datetime(2025,1,1).date(),
        max_value=datetime.now().date()
    )

    data_final = datetime.now().date()

    if st.button("🚀 Extrair e subir arquivo"):
        start_dt_str = data_inicial.strftime("%Y-%m-%dT00:00:00Z")
        end_dt_str = data_final.strftime("%Y-%m-%dT23:59:59Z")

        with st.spinner("Extraindo dados com paginação..."):
            tickets = get_all_tickets(start_dt_str, end_dt_str)
            st.write(f"Total tickets extraídos: {len(tickets)}")

            actions_data = extract_actions(tickets)
            df = pd.DataFrame(actions_data)

            # Aqui você pode aplicar tratamentos similares aos seus (exemplo)
            # Exemplo: converter datas, renomear colunas etc.

            csv_path = "tickets_movidesk_pag.csv"
            df.to_csv(csv_path, index=False)
            st.success(f"Arquivo salvo: {csv_path}")

            uploadSharePoint(csv_path, sharepoint_folder)

            st.dataframe(df.head())

        st.balloons()

if __name__ == "__main__":
    main()
