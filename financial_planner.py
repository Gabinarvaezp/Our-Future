import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.drawing.image import Image
from openpyxl.chart import PieChart, Reference, BarChart

class FinancialPlanner:
    def __init__(self):
        self.exchange_rate = 0.00025  # 1 COP = 0.00025 USD
        self.workbook = Workbook()
        self.setup_colors()
        
    def setup_colors(self):
        self.colors = {
            'pastel_blue': 'B4D8E7',
            'pastel_yellow': 'FFF4BD',
            'pastel_green': 'C8E6C9',
            'pastel_pink': 'FFD1DC'
        }

    def create_monthly_budget(self):
        ws = self.workbook.active
        ws.title = "Presupuesto Mensual"
        
        # Datos de ingresos
        income_usd = {
            'Jorge (USD)': 1506,
            'Gabby (COP convertido a USD)': 2170000 * self.exchange_rate
        }
        
        # Gastos Jorge (USD)
        expenses_jorge = {
            'SchoolFirst': 400,
            'Cozy House': 800,  # 400 * 2 por ser quincenal
            'Cooper': 730,
            'Car Insurance': 140,
            'Esposa': 100,
            'Golf': 100,
            'Salidas Amigos': 100
        }
        
        # Gastos Gabby (COP convertidos a USD)
        expenses_gabby = {
            'Casa': 400000 * self.exchange_rate,
            'Arreglo Casa': 800000 * self.exchange_rate,
            'Tarjeta Crédito': 500000 * self.exchange_rate,
            'Esposo': 200000 * self.exchange_rate,
            'Comida': 150000 * self.exchange_rate,
            'Transporte': 150000 * self.exchange_rate,
            'Cozy House': 400000 * self.exchange_rate
        }

        # Crear secciones
        self.create_section_header(ws, "INGRESOS MENSUALES", 1, self.colors['pastel_blue'])
        row = 2
        for source, amount in income_usd.items():
            ws.cell(row=row, column=1, value=source)
            ws.cell(row=row, column=2, value=amount)
            row += 1

        # Gastos Jorge
        row += 2
        self.create_section_header(ws, "GASTOS JORGE (USD)", row, self.colors['pastel_yellow'])
        row += 1
        for expense, amount in expenses_jorge.items():
            ws.cell(row=row, column=1, value=expense)
            ws.cell(row=row, column=2, value=amount)
            row += 1

        # Gastos Gabby
        row += 2
        self.create_section_header(ws, "GASTOS GABBY (USD)", row, self.colors['pastel_pink'])
        row += 1
        for expense, amount in expenses_gabby.items():
            ws.cell(row=row, column=1, value=expense)
            ws.cell(row=row, column=2, value=amount)
            row += 1

        # Crear gráficos
        self.create_pie_chart(ws, "Distribución de Gastos", row + 2)
        
    def create_debt_control(self):
        ws = self.workbook.create_sheet("Control de Deudas")
        # Implementar el control de deudas similar a la imagen
        
    def create_savings_tracker(self):
        ws = self.workbook.create_sheet("Tracking de Ahorro")
        
        # Metas de ahorro
        savings_goals = {
            'Viaje Cartagena': 1100,
            'Hotel All Inclusive': 650,
            'Vuelos': 200,
            'Transporte': 80,
            'Cena Romántica': 100
        }
        
        # Crear sección de metas
        self.create_section_header(ws, "METAS DE AHORRO", 1, self.colors['pastel_green'])
        row = 2
        for goal, amount in savings_goals.items():
            ws.cell(row=row, column=1, value=goal)
            ws.cell(row=row, column=2, value=amount)
            row += 1
            
        # Crear gráfico de progreso
        self.create_progress_chart(ws, row + 2)

    def create_section_header(self, ws, title, row, color):
        cell = ws.cell(row=row, column=1, value=title)
        cell.fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center")
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=2)

    def create_pie_chart(self, ws, title, row):
        pie = PieChart()
        pie.title = title
        # Configurar datos para el gráfico
        data = Reference(ws, min_col=2, min_row=2, max_row=15)
        labels = Reference(ws, min_col=1, min_row=2, max_row=15)
        pie.add_data(data)
        pie.set_categories(labels)
        ws.add_chart(pie, f"D{row}")

    def create_progress_chart(self, ws, row):
        chart = BarChart()
        chart.title = "Progreso de Ahorro"
        # Configurar datos para el gráfico
        data = Reference(ws, min_col=2, min_row=2, max_row=6)
        labels = Reference(ws, min_col=1, min_row=2, max_row=6)
        chart.add_data(data)
        chart.set_categories(labels)
        ws.add_chart(chart, f"D{row}")

    def save(self, filename="Financial_Planner.xlsx"):
        self.workbook.save(filename)

# Crear y guardar el archivo
planner = FinancialPlanner()
planner.create_monthly_budget()
planner.create_debt_control()
planner.create_savings_tracker()
planner.save()

print("¡Archivo Excel creado exitosamente!")
