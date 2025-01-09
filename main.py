import flet as ft
import pyodbc
from datetime import datetime
import locale
import time
import threading
import pandas as pd


data = []  
total_vendas_valor = 0  
locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')  
codigo = "000496"
data_atual = datetime.now()
data_init1 = data_atual.replace(day=1).strftime('%Y-%m-%d')
data_finish1 = data_atual.strftime('%Y-%m-%d')


def main(pagina:ft.Page):
    global table, fabric_table,loading_ring,codigo,data_atual,data_init1,data_finish1

    connection_string = (
        "Driver={SQL Server};"
        "Server=SRVSQL01;"
        "Database=MOINHO;"
        "UID=???;"  
        "PWD=????"  
        "Trusted_Connection=no;"
    )
    def saudacoes(): # Inserido uma conexão para retornar o nome perante ao codigo inserido no começo do código
            global nome
            conn = pyodbc.connect(connection_string)
            cursor = conn.cursor()
            cursor.execute(f"""
                            SELECT TOP 1 
                            nome 
                            FROM usuario 
                            WHERE cd_usuario = '{codigo}'
                        """)
            data = cursor.fetchall()
            
            if data:
                nome = data[0][0].split()[0]
            else:
                nome = "Nome não encontrado. 😒"

            cursor.close()
            conn.close()

            hora = int(datetime.now().strftime("%H"))
            
            if hora < 12:
                return ft.Text(f"Bom dia 🌞 {nome.split('-')[0]}",
                            font_family="Poppins3", size=20)
            elif hora < 18:
                return ft.Text(f"Boa Tarde ☀ {nome.split('-')[0]}",
                            font_family="Poppins3", size=20)
            else:
                return ft.Text(f"Boa Noite 🌃 {nome.split('-')[0]}",
                            font_family="Poppins3", size=20)

    saudacao_componente = saudacoes()
    pagina.controls.append(saudacao_componente)

    pagina.theme_mode = 'dark'
    pagina.window.maximized = True
    pagina.title = f'Estatísticas de Vendas 1.0 - {nome.split('-')[0]}'
    pagina.fonts = {
        "Poppins": "fonts/Poppins-Bold.ttf",
        "Poppins2": "fonts/Poppins-Light.ttf",
        "Poppins3": "fonts/Poppins-Regular.ttf",
    }

    loading_ring = ft.ProgressRing(visible=False)
    

    # Carregar metas do Excel
    def carregar_metas_excel():
        codigo_excel = codigo[3:] 
        excel = pd.read_excel(r'L:\Televendas\Sistema Televendas\Arqui\Metas.xls', sheet_name='Vendedores')
        excel['Vendedor'] = excel['Vendedor'].astype(str)
        # Retornando um dicionário de fabricantes e metas
        metas_filtradas1 = excel[excel['Vendedor'] == codigo_excel]
        return {row['Fabricante']: row['Valor'] for _, row in metas_filtradas1.iterrows()}


    def create_row(item):
                
        def open_modal(e):
            
            codigo_lançadora = item["Código Vendedor"]
            lancador = item["Lançador"]
            pedido = item["Pedido"]

            
            conn = pyodbc.connect(connection_string)
            cursor = conn.cursor()

           
            cursor.execute(f"""
                SELECT cd_vend, cd_vend_lc, nu_ped, valor_tot, dt_ped, qtde_volumes, cd_forn, situacao 
                FROM ped_vda 
                WHERE nu_ped = '{pedido}' and cd_vend_lc = '{lancador}'
            """)

            detalhes_pedido = cursor.fetchone()

            if detalhes_pedido:
                total_vendas_valor = detalhes_pedido[3]
                quantidade = detalhes_pedido[5]
                fornecedor = detalhes_pedido[6]
                situacao = detalhes_pedido[7]
                formatted_num = locale.format_string("%.2f", total_vendas_valor, grouping=True)
            else:
                formatted_num = "Nenhuma informação encontrada."

            cursor.close()
           
            modal_content = ft.Column(
                [
                    ft.Text(f"Código Vendedor: {codigo_lançadora}", font_family="Poppins3",size=15),
                    ft.Text(f"Lançador: {lancador}", font_family="Poppins3",size=15),
                    ft.Text(f"Pedido: {pedido}", font_family="Poppins3",size=15),
                    ft.Text(f"Quantidade: {int(quantidade)}", font_family="Poppins3",size=15),
                    ft.Text(f"Fornecedor: {fornecedor}", font_family="Poppins3",size=15),
                    ft.Text(f"Situação: {situacao}", font_family="Poppins3",size=15),
                    ft.Text(f"Total do Pedido: R$ {formatted_num}", font_family="Poppins3",size=15),
                ],
                tight=True
            )
            

            modal = ft.AlertDialog(
                title=ft.Text(f"📋 Detalhes do Pedido {pedido}", font_family="Poppins3"),
                content=modal_content,
                actions_alignment=ft.MainAxisAlignment.END,
            )

            pagina.dialog = modal
            modal.open = True
            pagina.update()

        return ft.DataRow(
            color="#000000",
            cells=[
                ft.DataCell(ft.Container(content=ft.Text(item["Pedido"], font_family="Poppins3",size=14), height=40)),
                ft.DataCell(ft.Container(content=ft.Text(item["Lançador"], font_family="Poppins3",size=14), height=40)),
                ft.DataCell(ft.Container(content=ft.Text(item["Código Vendedor"], font_family="Poppins3",size=14), height=40)),
                ft.DataCell(ft.Container(content=ft.Text(item["Venda Faturada"], font_family="Poppins3",size=14), height=40)),
                ft.DataCell(ft.Container(content=ft.Text(item["Periodo"], font_family="Poppins3",size=14), height=40)),
            ],
            on_select_changed=open_modal,
        )

    fabric_table = ft.DataTable(
        columns=[
            ft.DataColumn(label=ft.Row([ft.Text("Fabricante", font_family="Poppins", size=15), ft.Icon(name=ft.icons.ARROW_DROP_DOWN)]),
                          on_sort=lambda e: sort_table("fabric_table", "Fabricante", key_func=lambda row: row.cells[0].content.value)),
            ft.DataColumn(label=ft.Row([ft.Text("Total de Vendas", font_family="Poppins", size=15), ft.Icon(name=ft.icons.ARROW_DROP_DOWN)]),
                          on_sort=lambda e: sort_table("fabric_table", "Total de Vendas", key_func=lambda row: float(row.cells[1].content.value.replace('R$ ', '').replace('.', '').replace(',', '.')), numeric=True)),
            ft.DataColumn(label=ft.Row([ft.Text("Meta", font_family="Poppins", size=15)])), 
        ],
        rows=[],
        border=ft.border.all(2),
        data_row_color='#000000',
        heading_row_height=40,
        width=600,
    )

    table = ft.DataTable(
        columns=[
            ft.DataColumn(label=ft.Text("Pedido",font_family="Poppins",size=15)),
            ft.DataColumn(label=ft.Text("Código Lançador", font_family="Poppins",size=15)),
            ft.DataColumn(label=ft.Text("Código Vendedor", font_family="Poppins",size=15)),
            ft.DataColumn(label=ft.Text("Venda Faturada", font_family="Poppins",size=15)),
            ft.DataColumn(label=ft.Text("Periodo", font_family="Poppins",size=15)),
        ],
        rows=[],  
        border=ft.border.all(2),
        data_row_color='#000000',
        heading_row_height=40,
        width=1810,
        
    )
    # Texts de resultados visuais do aplicativo
    mix_toi = ft.Text("",font_family='Poppins3',size=20,visible=False)
    total_vendas = ft.Text("", visible=False, font_family="Poppins3",size=20)
    total_vendas_ao_todo = ft.Text("", visible=False, font_family="Poppins3",size=20)
    client = ft.Text("",visible=False,font_family="Poppins3",size=20)
    clock_text = ft.Text(value="", size=40, color="blue")


    def show_success_animation():
        confetti_container.visible = True
        pagina.update()
        threading.Timer(10, hide_confetti).start()
    confetti_container = ft.Container(
        content=ft.Image(
            src=".\confete.gif",
            fit=ft.ImageFit.COVER
        ),
        expand=True,
        alignment=ft.alignment.center,
        visible=False,
    )
    def hide_confetti():
        confetti_container.visible = False
        pagina.update()


    # Mix() retorna o total de Mix Comercial por vendedora
    def mix():
        codigo_excel = codigo[3:]
        excel_meta_total = pd.read_excel(r'L:\Televendas\Sistema Televendas\Arqui\Metas.xls', sheet_name='Meta')
        excel_meta_total['Vendedor'] = excel_meta_total['Vendedor'].astype(str)
        metas_filtradas = excel_meta_total[excel_meta_total['Vendedor'] == codigo_excel]['MIX TOI'].values
        if len(metas_filtradas) > 0:
            meta_value5 = metas_filtradas[0]
            if meta is not None:
                meta.value = f' {locale.format_string("%.2f", meta_value5, grouping=True)}'
        else:
            meta_value5 = 0
            if meta is not None:
                meta.value = '❌ Meta TOI não encontrada'
        conn = pyodbc.connect(connection_string)
        cursor = conn.cursor()
        cursor.execute(f"""
                       
                        SELECT  ip.cd_prod, SUM(ip.vl_base_comissao * ip.qtde) AS valor_total_vendas 
                        FROM 
                        ped_vda pv
                        LEFT JOIN 
                        it_pedv ip ON pv.nu_ped = ip.nu_ped AND pv.cd_emp = ip.cd_emp
                        INNER JOIN prod_mix_comercial pmc on ip.cd_prod = pmc.cd_prod
                        WHERE 
                        pv.cd_vend_lc = '{codigo}'
                        AND pv.dt_ped BETWEEN '{data_init1}' AND '{data_finish1}'
                        AND ip.situacao NOT IN ('DV', 'CA', 'AB')
						AND pv.origem_pedido = 'T'
						AND pv.tp_ped IN ('CV','CF','CL','CT','CP','CM','CS','FU','HT','VF','VD','VI','VE','VH','VL','VG','VK','PC')
                        group by ip.cd_prod

                    """)
        
        data = cursor.fetchall()

        if data: 
            total = sum(row[1]for row in data)
            totalmix = locale.format_string("%.2f", total, grouping=True)
            mix_toi.value= f"👑 MIX TOI: {totalmix} / {locale.format_string("%.2f", meta_value5, grouping=True)}"
            mix_toi.visible = True

        else:
            mix_toi.value = f"😢 Valor não encontrado."
            mix_toi.visible = True
        conn.close()

    # Carrega Fabricantes por Vendedora
    def carregar_fabricantes():

        conn = pyodbc.connect(connection_string)
        cursor = conn.cursor()
        cursor.execute(f"""
                SELECT 
                    DISTINCT f.descricao AS fabricante,
                    SUM(ip.qtde * ip.vl_base_comissao) AS valor_total_produto
                FROM 
                    ped_vda pv
                JOIN 
                    it_pedv ip ON pv.nu_ped = ip.nu_ped AND pv.cd_emp = ip.cd_emp
                JOIN 
                    produto pr ON ip.cd_prod = pr.cd_prod
                JOIN 
                    fabric f ON pr.cd_fabric = f.cd_fabric
                WHERE 
                    pv.cd_vend_lc = '{codigo}'
                    AND pv.dt_ped BETWEEN '{data_init1}' AND '{data_finish1}'
                    AND ip.situacao NOT IN ('DV', 'CA', 'AB')
                    AND pv.origem_pedido = 'T'
                    AND pv.tp_ped IN ('CV', 'CF', 'CL', 'CT', 'CP', 'CM', 'CS', 'FU', 'HT', 'VF', 'VD', 'VI', 'VE', 'VH', 'VL', 'VG', 'VK', 'PC')
                    AND f.ativo = '1'
                GROUP BY 
                    f.descricao
        """)

        data_fabricantes = cursor.fetchall()
        cursor.close()
        conn.close()

        # Carregando metas do Excel
        metas = carregar_metas_excel()
        
        fabricantes = [row[0] for row in data_fabricantes]
        valores = [float(row[1]) if row[1] is not None else 0.00 for row in data_fabricantes]
        
        fabric_table.rows = [
        ft.DataRow(
            cells=[
                ft.DataCell(ft.Text(fabricante, font_family="Poppins3", size=14)),
                ft.DataCell(ft.Text(f"R$ {locale.format_string('%.2f', valor, grouping=True)}", font_family="Poppins3", size=14)),
                ft.DataCell(ft.Text(f"R$ {locale.format_string('%.2f', metas.get(fabricante, 0), grouping=True)}", font_family="Poppins3", size=14)),  # Adicionando a coluna Meta
            ]
        )
        for fabricante, valor in zip(fabricantes, valores)
    ]

    pagina.update()

    if 'metas_filtradas' in globals():
        meta = ft.Text(f"", font_family="Poppins3",size=20)
    else:
        meta = ft.Text("✔ R$ 0.00", font_family="Poppins3", size=20)


    progress_bar = ft.ProgressBar(value=0, color='Green', width=400, height=15)
    percent_label = ft.Text("0.00%", size=30, font_family="Poppins3")   

    def total(progress_bar, percent_label):


       
        locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')

        global metas_filtradas, tot, metas_filtradas_mensal

        try:

            codigo_excel = codigo[3:]


            excel_meta_total = pd.read_excel(r'L:\Televendas\Sistema Televendas\Arqui\Metas.xls', sheet_name='Meta')
            excel_meta_mensal = pd.read_excel(r'L:\Televendas\Sistema Televendas\Arqui\Metas.xls', sheet_name='Total')
            


            excel_meta_total['Vendedor'] = excel_meta_total['Vendedor'].astype(str)
            excel_meta_mensal['CÓDIGO'] = excel_meta_mensal['CÓDIGO'].astype(str)


            metas_filtradas = excel_meta_total[excel_meta_total['Vendedor'] == codigo_excel]['Meta'].values
            if len(metas_filtradas) > 0:
                meta_value1 = metas_filtradas[0]
            else:
                meta_value1 = 0


            metas_filtradas = excel_meta_mensal[excel_meta_mensal['CÓDIGO'] == codigo_excel]['VENDA (LIQ)'].values
            if len(metas_filtradas) > 0:
                meta_value2 = metas_filtradas[0]
            else:
                meta_value2 = 0
            
            meta_int = float(meta_value2)


            if meta_int > 0:
                progresso = meta_int / meta_value1
                progress_bar.value = min(progresso, 1) 
                percent_label.value = f"{min(progresso, 1) * 100:.2f}%"
            else:
                progress_bar.value = 0
                percent_label.value = "0.00%"

            if total_vendas_ao_todo is not None:
                total_vendas_ao_todo.value = f"💰 Total Faturado: R$ {locale.format_string('%.2f', meta_int, grouping=True)} / Meta: R$ {locale.format_string('%.2f', meta_value1, grouping=True)}"
                total_vendas_ao_todo.visible = True

            if percent_label.value == '100.00%':
                show_success_animation()
            else:
                pass

        except Exception as e:
            print(f"Erro durante o processamento: {e}")


    def clientes():
        conn = pyodbc.connect(connection_string)
        cursor = conn.cursor()
        cursor.execute(f"""SELECT
                         cd_clien 
                         FROM ped_vda 
                         WHERE cd_vend_lc = '000496' 
                         AND dt_ped BETWEEN '{data_init1}' AND '{data_finish1}' 
                         AND InicioProcessoFatura = '1' 
                        AND origem_pedido = 'T'
                        AND tp_ped IN ('CV', 'CF', 'CL', 'CT', 'CP', 'CM', 'CS', 'FU', 'HT', 'VF', 'VD', 'VI', 'VE', 'VH', 'VL', 'VG', 'VK', 'PC')
                        AND situacao NOT IN ('CA', 'DV')
                         group by cd_clien
                        """)
        
        cont = cursor.fetchall()
        total_client = len(cont)

        codigo_excel = codigo[3:]
        excel_meta_total = pd.read_excel(r'L:\Televendas\Sistema Televendas\Arqui\Metas.xls', sheet_name='Meta')
        excel_meta_total['Vendedor'] = excel_meta_total['Vendedor'].astype(str)
        metas_filtradas = excel_meta_total[excel_meta_total['Vendedor'] == codigo_excel]['POSITIVAÇÃO'].values
        if len(metas_filtradas) > 0:
            meta_value2 = metas_filtradas[0]
            if meta is not None:
                client.value = f"💹 Positivação: {total_client} / {locale.format_string("%.0f", meta_value2, grouping=True)}"
        else:
            meta_value2 = 0
            if meta is not None:
                client.value = f"💹 Positivação: {total_client} / ❌ Meta não encontrada "
        

        client.value = f"💹 Positivação: {total_client} / {locale.format_string("%.0f", meta_value2, grouping=True)}"
        client.visible = True

    # Inicia Carregamento de todas as Funções e animações de carregamento
    def iniciar_carregamento():
        loading_ring.visible = True
        pagina.update()
        time.sleep(0.8)

        def executar_carregamento():
            saudacoes()
            mix()
            carregar_fabricantes()
            total(progress_bar, percent_label)
            clientes()
            carregar_metas_excel()
            loading_ring.visible = False
            pagina.controls.append(saudacao_componente)
            pagina.update()

        threading.Thread(target=executar_carregamento).start()

    confetti_container = ft.Container(
    content=ft.Stack(
        [
            ft.Image(
                src="L:\Televendas\Sistema Televendas\Arqui\confete.gif",
                fit=ft.ImageFit.COVER,  
                width=2000,
                height=1000,
            ),
            ft.Text(
                value=f'🏆 PARABÉNS {nome.split('-')[0]}, SUA META FOI BATIDA !! 🏆',
                font_family="Poppins",
                size=75,
                color="Yellow",
                text_align=ft.TextAlign.CENTER,
                width=800, 
                height=500, 
            ),
        ],
        alignment=ft.alignment.center,
    ),
    alignment=ft.alignment.center, 
    visible=False, 
    )
    
    loading_ring = ft.ProgressRing(width=30, height=30, visible=True)
    
    botao = ft.ElevatedButton("Atualizar",
                               icon=ft.icons.REFRESH,
                                 on_click=lambda _: iniciar_carregamento(),
                                 width=200,
                                 height=35,
                                 bgcolor=ft.colors.GREEN,
                                 color='White',
                                 style=ft.ButtonStyle(
                                     text_style=ft.Text(
                                         font_family='Poppins',
                                         size=50)
                                         )
                                 )

    scrollable_column = ft.Column(
        controls=[fabric_table],
        scroll=ft.ScrollMode.AUTO,
        height=720,
    )

    layout = ft.Container(
        content=ft.Row(
            [
                ft.Container(
                    content=scrollable_column,
                ),
                ft.Container(
                    content=ft.Column(
                        [
                            loading_ring,
                            percent_label,
                            progress_bar,
                            total_vendas,
                            total_vendas_ao_todo,
                            client,
                            mix_toi,
                            botao,
                            clock_text,
                            confetti_container
                        ],
                        alignment=ft.MainAxisAlignment.CENTER,  
                        horizontal_alignment=ft.CrossAxisAlignment.CENTER,  
                        spacing=25,
                    ),
                    expand=True,  
                    alignment=ft.alignment.center,  
                ),
            ],
        ),
        padding=40,
        margin=5,
    )
    
    def sort_table(table_name,colum_name, key_func=None, numeric=False):
        global table, fabric_table
        table_to_sort = table if table_name == "table" else fabric_table
        table_to_sort.rows.sort(key=key_func, reverse=True if numeric else False)
        pagina.update()

    pagina.add(
        ft.Stack(
            [
                layout,
                confetti_container
            ]
        )
    )
    pagina.overlay.append(confetti_container)

    iniciar_carregamento()
ft.app(target=main)
