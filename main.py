import flet as ft
import pyodbc
from datetime import datetime
import locale
import time
import threading
import matplotlib.pyplot as plt
from io import BytesIO

data = []  
total_vendas_valor = 0  
locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')  

def main(pagina: ft.Page):
    global table
    pagina.theme_mode = 'dark'
    pagina.window_maximized = True
    pagina.title = 'Estatisticas de Vendas - Gessica'
    pagina.fonts = {
        "Poppins": "fonts/Poppins-Bold.ttf",
        "Poppins2": "fonts/Poppins-Light.ttf",
        "Poppins3": "fonts/Poppins-Regular.ttf",
    }
    def grafico(e):
        conn = pyodbc.connect(connection_string)
        cursor = conn.cursor()
        
        # Parâmetros
        codigo = "000400"
        data_atual = datetime.now()
        data_init1 = data_atual.replace(day=1).strftime('%Y-%m-%d')  # Início do mês atual
        data_finish1 = data_atual.strftime('%Y-%m-%d')  # Data atual

        # Executando a query
        cursor.execute(f"""
            SELECT 
                f.descricao AS fabricante,
                SUM(pv.valor_tot) AS valor_total_vendas
            FROM 
                ped_vda pv
            JOIN 
                it_pedv ip ON pv.nu_ped = ip.nu_ped
            JOIN 
                produto pr ON ip.cd_prod = pr.cd_prod
            JOIN 
                fabric f ON pr.cd_fabric = f.cd_fabric
            WHERE 
                pv.cd_vend_lc = '{codigo}'
                AND pv.dt_ped BETWEEN '{data_init1}' AND '{data_finish1}'
            GROUP BY 
                f.descricao;
        """)
        
        # Obtendo os resultados
        data = cursor.fetchall()
        
        # Fechar o cursor e a conexão
        cursor.close()
        conn.close()

        # Separando fabricantes e valores em listas
        fabricantes = [row[0] for row in data]  # Descrição dos fabricantes
        valores = [row[1] for row in data]  # Valores de vendas

        # Verificando o resultado (opcional)
        print(fabricantes, valores)

        # Criando o gráfico
        plt.bar(fabricantes, valores)
        plt.xlabel('Fabricantes')
        plt.ylabel('Total de Vendas')
        plt.title('Vendas por Fabricante')
        plt.xticks(rotation=45, ha='right')  # Rotaciona os rótulos para melhor visualização
        plt.show()
        # Salvar o gráfico em um objeto BytesIO
        buf = BytesIO()
        plt.savefig(buf, format='png')
        plt.close()  # Fecha o gráfico para liberar a memória
        buf.seek(0)  # Retorna o ponteiro ao início do objeto BytesIO
        
        # Cria um componente de imagem Flet
        img = ft.Image(src=buf.getvalue(), width=800, height=400)

        # Criar um modal para exibir o gráfico
        modal = ft.AlertDialog(
            title="Gráfico de Vendas por Fabricante",
            content=img,
            actions=[ft.ElevatedButton("Fechar", on_click=lambda e: e.page.dialog.close())],
        )
        
        # Exibe o modal
        ft.App.get_current().dialog = modal
        modal.open()
        ft.App.get_current().update()

    global loading_ring
    loading_ring = ft.ProgressRing(visible=False)
    # Função para filtrar os itens da tabela com base na pesquisa
    def filtrar_tabela(termo):
        termo = termo.lower()  # Garantir que o termo de pesquisa esteja em letras minúsculas
        termos_filtrados = [
            item for item in data
            if termo in str(item["Código Vendedor"]).lower()
            or termo in str(item["Lançador"]).lower()
            or termo in str(item["Pedido"]).lower()
            or termo in str(item["Venda Faturada"]).lower()  # Converte para string
            or termo in str(item["Periodo"]).lower()
        ]
        table.rows = [create_row(item) for item in termos_filtrados]
        pagina.update()

    
    def saudacoes():
    # Saudações
        
        hora = int(datetime.now().strftime("%H"))
        if hora < 12:
            return ft.Text("Bom dia 🌞 Gessica",
                                font_family="Poppins3", size=20)
        elif hora < 18:
            return ft.Text("Boa Tarde ☀ Gessica",
                                font_family="Poppins3",size=20)
        else:
            return ft.Text("Boa Noite 🌃 Gessica",
                                font_family="Poppins3",size=20)
        
    saudacao_componente = saudacoes()
    pagina.controls.append(saudacao_componente)

    connection_string = (
        "Driver={SQL Server};"
        "Server=SRVSQL01;"
        "Database=MOINHO;"
        "UID=TargetAdmin;"  
        "PWD=dlh%9>?xiyh1QPB;"  
        "Trusted_Connection=no;"
    )

    def create_row(item):
                
        def open_modal(e):
            
            codigo_vendedor = item["Código Vendedor"]
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
                    ft.Text(f"Código Vendedor: {codigo_vendedor}", font_family="Poppins3",size=15),
                    ft.Text(f"Lançador: {lancador}", font_family="Poppins3",size=15),
                    ft.Text(f"Pedido: {pedido}", font_family="Poppins3",size=15),
                    ft.Text(f"Quantidade: {int(quantidade)}", font_family="Poppins3",size=15),
                    ft.Text(f"Fornecedor: {fornecedor}", font_family="Poppins3",size=15),
                    ft.Text(f"Situação: {situacao}", font_family="Poppins3",size=15),
                    ft.Text(f"Total do Pedido: R$ {formatted_num}", font_family="Poppins3",size=15),
                ],
                tight=True
            )
            
            # Configura o modal
            modal = ft.AlertDialog(
                title=ft.Text(f"📋 Detalhes do Pedido {pedido}", font_family="Poppins3"),
                content=modal_content,
                actions_alignment=ft.MainAxisAlignment.END,
            )
    
            # Exibir o modal
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
    progress = ft.ProgressBar(value=30/100, 
                              bar_height=10,
                              width=500,
                              bgcolor='white',
                              color='blue',
                              border_radius=30
                              )
    cont_ped_title = ft.Text("Pedidos:",font_family='Poppins3',size=20)
    cont_ped = ft.Text("",font_family='Poppins3',size=20)
    # Adicionar campo de pesquisa
    campo_pesquisa = ft.TextField(
        label="Pesquisar",
        hint_text="Digite para pesquisar...",
        on_change=lambda e: filtrar_tabela(e.control.value),  # Chamar a função ao mudar o texto
        color="#F0F0F0"
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

    total_vendas = ft.Text("", visible=False, font_family="Poppins3",size=20)
    total_vendas_ao_todo = ft.Text("", visible=False, font_family="Poppins3",size=20)

    def total():
        codigo = "000400"
        conn = pyodbc.connect(connection_string)
        cursor = conn.cursor()

        data_atual = datetime.now()
        data_init1 = data_atual.replace(day=1).strftime('%Y-%m-%d')
        data_finish1 = data_atual.strftime('%Y-%m-%d')

        cursor.execute(f"SELECT cd_vend, cd_vend_lc, nu_ped, valor_tot, dt_ped,InicioProcessoFatura FROM ped_vda WHERE cd_vend_lc = '{codigo}' AND dt_ped BETWEEN '{data_init1}' AND '{data_finish1}' and InicioProcessoFatura = '1' AND cd_emp = '3'")
        vendas = cursor.fetchall()
        total_vendas_valor = sum(row[3] for row in vendas)
        formatted_num = locale.format_string("%.2f", total_vendas_valor, grouping=True)
        print(formatted_num , total_vendas_valor)
        total_vendas_ao_todo.value = f"Total Faturado esse Mês: 💰 {formatted_num}  / "
        total_vendas_ao_todo.visible = True
        cursor.close()
    total()
    
    def pesquisar():
        codigo = "000400"
        conn = pyodbc.connect(connection_string)
        cursor = conn.cursor()

        data_atual = datetime.now()
        data_init1 = data_atual.replace(day=1).strftime('%Y-%m-%d')
        data_finish1 = data_atual.strftime('%Y-%m-%d')

        cursor.execute(f"SELECT cd_vend, cd_vend_lc, nu_ped, valor_tot, dt_ped,InicioProcessoFatura FROM ped_vda WHERE cd_vend_lc = '{codigo}' AND dt_ped BETWEEN '{data_init1}' AND '{data_finish1}' and InicioProcessoFatura = '1' AND cd_emp = '3'")

        global data 
        data = [
            { "Pedido": str(row.nu_ped),"Lançador": str(row.cd_vend_lc), "Código Vendedor": str(row.cd_vend), "Venda Faturada":locale.format_string("%.2f", int(row.valor_tot), grouping=True), "Periodo": str(row.dt_ped)}
            for row in cursor.fetchall()
        ]

        total_ped = len(data)
        cont_ped.value = f"{total_ped}"
        cont_ped.visible = True

        table.rows = [create_row(item) for item in data]
        pagina.update()
        cursor.close()

    def iniciar_carregamento():
        loading_ring.visible = True
        pagina.update()
        time.sleep(0.8)

        # Executa a função pesquisar em uma nova thread
        def executar_pesquisa():
            pesquisar()
            loading_ring.visible = False
            pagina.update()

        threading.Thread(target=executar_pesquisa).start()
    saudacoes()
    # Criando o indicador de carregamento (ProgressRing)
    loading_ring = ft.ProgressRing(width=30, height=30, visible=True)
    
    botao = ft.ElevatedButton("Atualizar",
                               icon=ft.icons.REFRESH,
                                 on_click=lambda _: iniciar_carregamento(),
                                 width=200,
                                 height=35
                                 )
    detalhes = ft.ElevatedButton("Gráfico",
                                  icon=ft.icons.DASHBOARD_CUSTOMIZE,
                                  width=200,
                                  height=35,
                                  on_click=grafico
                                  )
    
    # Coluna com scroll
    scrollable_column = ft.Column(
        controls=[ table],
        scroll=ft.ScrollMode.AUTO,
        width=1850,
        height=720,
    )

    layout = ft.Container(
        content=ft.Column(
            [
                ft.Row([loading_ring,total_vendas, total_vendas_ao_todo,cont_ped_title,cont_ped, botao, detalhes,progress]),
                ft.ResponsiveRow([campo_pesquisa]),
                ft.ResponsiveRow([scrollable_column]),
            ],
            spacing=25,
        ),
        padding=40,
        margin=5,
    )

    pagina.add(layout)

    iniciar_carregamento()

ft.app(target=main)
