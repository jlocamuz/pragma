"""
Generador de archivos Excel usando Pandas
Solo muestra los datos calculados por hours_calculator
"""

import os
import pandas as pd
from datetime import datetime
from typing import Dict
from config.default_config import DEFAULT_CONFIG
import re 


class ExcelReportGenerator:
    def __init__(self):
        self.output_dir = os.path.expanduser(DEFAULT_CONFIG['output_directory'])
        self.filename_format = DEFAULT_CONFIG['filename_format']

    # -------------------- Helpers --------------------
    def hours_to_excel_time(self, hours: float) -> float:
        """
        Excel guarda tiempos como fracción del día.
        1 hora = 1/24 ≈ 0.041666..., 1.5 h -> 0.0625, etc.
        """
        try:
            return round((float(hours) if hours else 0.0) / 24.0, 10)
        except Exception:
            return 0.0

    def _only_hhmm(self, value) -> str:
        """Devuelve 'HH:MM' si lo encuentra dentro de value; si no, ''."""
        if not value:
            return ""
        m = re.search(r'([01]\d|2[0-3]):[0-5]\d', str(value))
        return m.group(0) if m else ""

    def generate_report(self, processed_data: Dict, start_date: str, end_date: str, output_filename: str = None) -> str:
        """Genera el reporte Excel usando pandas"""

        # Preparar datos para cada hoja
        summary_data = self._prepare_summary_data(processed_data)
        daily_data = self._prepare_daily_data(processed_data)
        config_data = self._prepare_config_data()

        # Nombre de archivo
        if not output_filename:
            output_filename = self.filename_format.format(
                start_date=start_date.replace('-', ''), end_date=end_date.replace('-', '')
            )

        os.makedirs(self.output_dir, exist_ok=True)
        filepath = os.path.join(self.output_dir, output_filename)

        # Escribir a Excel con formato
        with pd.ExcelWriter(filepath, engine='xlsxwriter') as writer:
            # Hoja Resumen
            summary_df = pd.DataFrame(summary_data)
            summary_df.to_excel(writer, sheet_name='Resumen Consolidado', index=False, startrow=3)
            self._format_summary_sheet(writer, summary_df, start_date, end_date)

            # Hoja Detalle Diario
            daily_df = pd.DataFrame(daily_data)
            daily_df.to_excel(writer, sheet_name='Detalle Diario', index=False, startrow=3)
            self._format_daily_sheet(writer, daily_df, start_date, end_date)

            # Hoja Configuración
            config_df = pd.DataFrame(config_data)
            config_df.to_excel(writer, sheet_name='Configuración', index=False, startrow=4)
            self._format_config_sheet(writer, config_df, start_date, end_date)

        print(f"✅ Reporte Excel generado: {filepath}")
        return filepath

    # -------------------- Preparación de datos --------------------
    def _prepare_summary_data(self, processed_data: Dict) -> list:
        """Prepara datos para la hoja de resumen (horas como tiempo Excel con hh:mm)"""
        summary_rows = []

        for emp in processed_data.values():
            info = emp['employee_info']
            totals = emp['totals']

            row = {
                'ID Empleado': info.get('employeeInternalId', ''),
                'Nombre': info.get('firstName', ''),
                'Apellido': info.get('lastName', ''),
                'Total Horas': self.hours_to_excel_time(totals.get('total_hours_worked', 0.0)),
                'Horas Regulares': self.hours_to_excel_time(totals.get('total_regular_hours', 0.0)),
                'Horas Extra 50%': self.hours_to_excel_time(totals.get('total_extra_hours_50', 0.0)),
                'Horas Extra 100%': self.hours_to_excel_time(totals.get('total_extra_hours_100', 0.0)),
                'Horas Nocturnas': self.hours_to_excel_time(totals.get('total_night_hours', 0.0)),
                'Horas Feriado': self.hours_to_excel_time(totals.get('total_holiday_hours', 0.0)),
                'Total Tardanzas': self.hours_to_excel_time(totals.get('total_tardanza_horas', 0.0)),
                'Total Retiros Anticipados': self.hours_to_excel_time(totals.get('total_retiro_anticipado_horas', 0.0)),
                'Horas Extra Diurnas': self.hours_to_excel_time(totals.get('total_extra_day_hours', 0.0)),
                'Horas Extra Nocturnas': self.hours_to_excel_time(totals.get('total_extra_night_hours', 0.0)),
            }
            summary_rows.append(row)

        return summary_rows

    def _prepare_daily_data(self, processed_data: Dict) -> list:
        """Prepara datos para la hoja de detalle diario - Solo muestra los datos ya calculados"""
        daily_rows = []

        for emp in processed_data.values():
            info = emp['employee_info']

            for d in emp['daily_data']:
                # Observaciones
                observations = []

                if d.get('is_holiday'):
                    observations.append(f"Feriado: {d.get('holiday_name') or 'N/A'}")
                if d.get('has_time_off'):
                    observations.append(f"Licencia: {d.get('time_off_name') or 'N/A'}")
                if d.get('has_absence'):
                    observations.append("AUSENCIA SIN AVISO")

                row = {
                    'Legajo': info.get('employeeInternalId', ''),
                    'Apellido, Nombre': info.get('lastName', '') + ', ' + info.get('firstName', ''),
                    'Fecha': d.get('day_of_week', '') + ' ' + d.get('date', ''),
                    'Horario obligatorio': d.get('time_range'),
                    'Fichadas': self._only_hhmm(d.get('shift_start', '')) + ' - ' + self._only_hhmm(d.get('shift_end', '')),
                    'Observaciones': ', '.join(observations) if observations else '',

                    # Horas en formato Excel (ya calculadas en hours_calculator)
                    'Horas Trabajadas': self.hours_to_excel_time(d.get('hours_worked', 0.0)),
                    'Horas extra': self.hours_to_excel_time(d.get('extra_hours', 0.0)),
                    'Horas Extra Diurnas': self.hours_to_excel_time(d.get('extra_hours_day', 0.0)),
                    'Horas Extra Nocturnas': self.hours_to_excel_time(d.get('extra_hours_night', 0.0)),
                    'Horas Regulares': self.hours_to_excel_time(d.get('regular_hours', 0.0)),
                    'Horas Extra 50%': self.hours_to_excel_time(d.get('extra_hours_50', 0.0)),
                    'Horas Extra 100%': self.hours_to_excel_time(d.get('extra_hours_100', 0.0)),
                    'Horas Extra 150%': self.hours_to_excel_time(d.get('extra_hours_150', 0.0)),
                    'Horas Nocturnas': self.hours_to_excel_time(d.get('night_hours', 0.0)),
                    'Horas Feriado': self.hours_to_excel_time(d.get('holiday_hours', 0.0)),
                    'Es Franco': 'Sí' if d.get('is_rest_day') else 'No',
                    'Es Feriado': 'Sí' if d.get('is_holiday') else 'No',
                    'Nombre Feriado': d.get('holiday_name') or '',
                    'Tiene Licencia': 'Sí' if d.get('has_time_off') else 'No',
                    'Tipo Licencia': d.get('time_off_name') or '',
                    
                    # Tardanza y Retiro Anticipado YA calculados en hours_calculator
                    'Tardanza': self.hours_to_excel_time(d.get('tardanza_horas', 0.0)),
                    'Retiro Anticipado': self.hours_to_excel_time(d.get('retiro_anticipado_horas', 0.0)),
                }
                daily_rows.append(row)

        return daily_rows

    def _prepare_config_data(self) -> list:
        """Prepara datos de configuración"""
        return [
            {'Clave': 'output_directory', 'Valor': DEFAULT_CONFIG.get('output_directory', '')},
            {'Clave': 'filename_format', 'Valor': DEFAULT_CONFIG.get('filename_format', '')},
            {'Clave': 'extras_al_50', 'Valor': DEFAULT_CONFIG.get("extras_al_50", 2)},
            {'Clave': 'hora_nocturna_inicio', 'Valor': DEFAULT_CONFIG.get('hora_nocturna_inicio', 21)},
            {'Clave': 'hora_nocturna_fin', 'Valor': DEFAULT_CONFIG.get('hora_nocturna_fin', 6)},
            {'Clave': 'sabado_limite_hora', 'Valor': DEFAULT_CONFIG.get('sabado_limite_hora', 13)},
            {'Clave': 'local_timezone', 'Valor': DEFAULT_CONFIG.get('local_timezone', 'America/Argentina/Buenos_Aires')},
        ]

    # -------------------- Formato de hojas --------------------
    def _format_summary_sheet(self, writer, df, start_date, end_date):
        """Aplica formato a la hoja de resumen"""
        workbook = writer.book
        worksheet = writer.sheets['Resumen Consolidado']

        # Formatos base
        header_format = workbook.add_format({
            'bold': True, 'font_color': 'white', 'bg_color': '#366092',
            'border': 1, 'align': 'center', 'valign': 'vcenter'
        })
        title_format = workbook.add_format({'bold': True, 'font_size': 12})
        time_format = workbook.add_format({'num_format': 'hh:mm'})

        # Colores de columnas (opcionales)
        regular_format = workbook.add_format({'bg_color': '#D4EDDA', 'border': 1, 'num_format': 'hh:mm'})
        extra_50_format = workbook.add_format({'bg_color': '#FFF3CD', 'border': 1, 'num_format': 'hh:mm'})
        extra_100_format = workbook.add_format({'bg_color': '#F8D7DA', 'border': 1, 'num_format': 'hh:mm'})
        night_format = workbook.add_format({'bg_color': '#D1ECF1', 'border': 1, 'num_format': 'hh:mm'})
        holiday_format = workbook.add_format({'bg_color': '#D6EAF8', 'border': 1, 'num_format': 'hh:mm'})

        # Títulos
        worksheet.write(0, 0, "REPORTE DE ASISTENCIA - RESUMEN CONSOLIDADO", title_format)
        worksheet.write(1, 0, f"Período: {start_date} al {end_date}", title_format)
        worksheet.write(2, 0, f"Generado: {datetime.now().strftime('%d/%m/%Y %H:%M')}", title_format)

        # Columnas que deben mostrarse como tiempo
        time_cols = {
            'Total Horas', 'Horas Regulares', 'Horas Extra 50%', 'Horas Extra 100%', 
            'Horas Nocturnas', 'Horas Feriado', 'Total Tardanzas', 'Total Retiros Anticipados',
            'Horas Extra Diurnas', 'Horas Extra Nocturnas',
        }

        # Encabezados (fila 3)
        for col_num, col_name in enumerate(df.columns):
            worksheet.write(3, col_num, col_name, header_format)

            # Anchos y formatos
            if col_name in time_cols:
                # Formato general hh:mm (y color específico si aplica)
                if col_name == 'Horas Regulares':
                    worksheet.set_column(col_num, col_num, 16, regular_format)
                elif col_name == 'Horas Extra 50%':
                    worksheet.set_column(col_num, col_num, 16, extra_50_format)
                elif col_name == 'Horas Extra 100%':
                    worksheet.set_column(col_num, col_num, 16, extra_100_format)
                elif col_name == 'Horas Nocturnas':
                    worksheet.set_column(col_num, col_num, 16, night_format)
                elif col_name == 'Horas Feriado':
                    worksheet.set_column(col_num, col_num, 16, holiday_format)
                else:
                    worksheet.set_column(col_num, col_num, 16, time_format)
            else:
                worksheet.set_column(col_num, col_num, 18)

    def _format_daily_sheet(self, writer, df, start_date, end_date):
        """Aplica formato a la hoja de detalle diario"""
        workbook = writer.book
        worksheet = writer.sheets['Detalle Diario']

        title_format = workbook.add_format({'bold': True, 'font_size': 12})
        header_format = workbook.add_format({
            'bold': True, 'font_color': 'white', 'bg_color': '#366092',
            'border': 1, 'align': 'center', 'valign': 'vcenter'
        })
        time_format = workbook.add_format({'num_format': 'hh:mm'})

        # Títulos
        worksheet.write(0, 0, "DETALLE DIARIO DE ASISTENCIA", title_format)
        worksheet.write(1, 0, f"Período: {start_date} al {end_date}", title_format)
        worksheet.write(2, 0, f"Generado: {datetime.now().strftime('%d/%m/%Y %H:%M')}", title_format)

        # Columnas con tiempos
        time_cols = {
            'Horas Trabajadas', 'Horas Regulares', 'Horas Extra 50%',
            'Horas Extra 100%', 'Horas Nocturnas', 'Horas Feriado', 
            'Horas extra', 'Tardanza', 'Retiro Anticipado',
            'Horas Extra Diurnas', 'Horas Extra Nocturnas',
        }

        # Encabezados (fila 3)
        for col_num, col_name in enumerate(df.columns):
            worksheet.write(3, col_num, col_name, header_format)

            if col_name in time_cols:
                worksheet.set_column(col_num, col_num, 12, time_format)
            elif col_name in ['Fecha', 'Nombre Feriado', 'Observaciones']: 
                worksheet.set_column(col_num, col_num, 20)
            elif col_name in ['Apellido, Nombre']:
                worksheet.set_column(col_num, col_num, 28)
            else:
                worksheet.set_column(col_num, col_num, 14)

    def _format_config_sheet(self, writer, df, start_date, end_date):
        """Aplica formato a la hoja de configuración"""
        workbook = writer.book
        worksheet = writer.sheets['Configuración']

        title_format = workbook.add_format({'bold': True, 'font_size': 12})

        # Títulos
        worksheet.write(0, 0, "CONFIGURACIÓN DEL SISTEMA", title_format)
        worksheet.write(1, 0, "Parámetros utilizados para el reporte", title_format)
        worksheet.write(2, 0, f"Período: {start_date} al {end_date}", title_format)
        worksheet.write(3, 0, f"Generado: {datetime.now().strftime('%d/%m/%Y %H:%M')}", title_format)

        # Ajustar anchos
        worksheet.set_column(0, 0, 28)  # Clave
        worksheet.set_column(1, 1, 55)  # Valor
