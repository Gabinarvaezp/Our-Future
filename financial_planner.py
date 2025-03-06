import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.drawing.image import Image
from openpyxl.chart import PieChart, Reference, BarChart
from datetime import datetime

class FinancialPlanner:
    def __init__(self):
        self.exchange_rate = 0.00025  # 1 COP = 0.00025 USD
        self.workbook = Workbook()
        self.setup_colors()
        
    def setup_colors(self):
        self.colors = {
            'header_blue': '4472C4',
            'header_yellow': 'FFD700',
            'pastel_blue': 'B4D8E7',
            'pastel_yellow': 'FFF4BD',
            'pastel_green': 'C8E6C9',
            'pastel_pink': 'FFD1DC'
        }

    def create_monthly_budget(self):
        ws = self.workbook.active
        ws.title = "Presupuesto Mensual"
        
        # Configurar ancho de columnas
        ws.column_dimensions['A'].width = 25
        ws.column_dimensions['B'].width = 15
        
        # Título principal
        self.create_fancy_header(ws, "PRESUPUESTO MENSUAL FAMILIAR", 1)
        
        # Primera Quincena
        row = 3
        self.create_section_header(ws, "PRIMERA QUINCENA", row, self.colors['header_blue'])
        row += 1
        
        # Ingresos Jorge
        ws.cell(row=row, column=1, value="Ingresos Jorge (USD)")
        ws.cell(row=row, column=2, value=1506)
        row += 1
        
        # Ingresos Gabby
        ws.cell(row=row, column=1, value="Ingresos Gabby (COP)")
        ws.cell(row=row, column=2, value=2170000)
        row += 2
        
        # Gastos Jorge
        self.create_section_header(ws, "GASTOS JORGE (USD)", row, self.colors['pastel_yellow'])
        row += 1
        expenses_jorge = {
            'SchoolFirst': 400,
            'Cozy House': 800,
            'Cooper': 730,
            'Car Insurance': 140,
            'Esposa': 100,
            'Golf': 100,
            'Salidas Amigos': 100
        }
        
        for expense, amount in expenses_jorge.items():
            ws.cell(row=row, column=1, value=expense)
            ws.cell(row=row, column=2, value=amount)
            row += 1
            
        # Gastos Gabby
        row += 1
        self.create_section_header(ws, "GASTOS GABBY (COP)", row, self.colors['pastel_pink'])
        row += 1
        expenses_gabby = {
            'Casa': 400000,
            'Arreglo Casa': 800000,
            'Tarjeta Crédito': 500000,
            'Esposo': 200000,
            'Comida': 150000,
            'Transporte': 150000,
            'Cozy House': 400000
        }
        
        for expense, amount in expenses_gabby.items():
            ws.cell(row=row, column=1, value=expense)
            ws.cell(row=row, column=2, value=amount)
            row += 1
            
        # Resumen y Gráficos
        self.create_monthly_summary(ws, row + 2)
        self.create_monthly_charts(ws, row + 10)

    def create_monthly_summary(self, ws, row):
        self.create_section_header(ws, "RESUMEN MENSUAL", row, self.colors['header_blue'])
        row += 1
        
        # Calcular totales
        total_income_usd = 1506 + (2170000 * self.exchange_rate)
        total_expenses_jorge = 2370  # Suma de gastos de Jorge
        total_expenses_gabby_usd = 2600000 * self.exchange_rate  # Suma de gastos de Gabby en USD
        
        summary = {
            'Total Ingresos (USD)': total_income_usd,
            'Total Gastos (USD)': total_expenses_jorge + total_expenses_gabby_usd,
            'Ahorro (USD)': total_income_usd - (total_expenses_jorge + total_expenses_gabby_usd)
        }
        
        for item, amount in summary.items():
            ws.cell(row=row, column=1, value=item)
            ws.cell(row=row, column=2, value=amount)
            row += 1

    def create_monthly_charts(self, ws, row):
        # Gráfico de Gastos
        pie = PieChart()
        pie.title = "Distribución de Gastos"
        data = Reference(ws, min_col=2, min_row=6, max_row=20)
        labels = Reference(ws, min_col=1, min_row=6, max_row=20)
        pie.add_data(data)
        pie.set_categories(labels)
        ws.add_chart(pie, f"D{row}")

    def create_fancy_header(self, ws, title, row):
        cell = ws.cell(row=row, column=1, value=title)
        cell.font = Font(size=16, bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color=self.colors['header_blue'], 
                              end_color=self.colors['header_blue'], 
                              fill_type="solid")
        cell.alignment = Alignment(horizontal="center")
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=2)

    def create_section_header(self, ws, title, row, color):
        cell = ws.cell(row=row, column=1, value=title)
        cell.fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center")
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=2)

    def save(self, filename=None):
        if filename is None:
            filename = f"Presupuesto_Mensual_{datetime.now().strftime('%Y%m%d')}.xlsx"
        self.workbook.save(filename)

# Crear y guardar el archivo
planner = FinancialPlanner()
planner.create_monthly_budget()
planner.save()

print("¡Archivo Excel creado exitosamente!")
