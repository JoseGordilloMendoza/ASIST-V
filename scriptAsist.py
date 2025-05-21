import csv
import pandas as pd
from datetime import datetime, timedelta
import argparse
import re
import os

def convert_duration_to_hours(duration_str):
    """Convierte la duración del formato 'X h Y min' a horas totales como flotante."""
    if pd.isna(duration_str):
        return 0.0
    
    parts = duration_str.split(' ')
    hours = 0.0
    minutes = 0.0
    
    if len(parts) >= 2 and 'h' in parts[1]:
        hours = float(parts[0])
    
    if len(parts) >= 4 and 'min' in parts[3]:
        minutes = float(parts[2])
        
    return hours + (minutes / 60)

def normalize_name(name):
    """Normaliza los nombres para facilitar la comparación."""
    if not name:
        return ""
    # Eliminar acentos y convertir a mayúsculas
    name = name.strip().upper()
    # Eliminar espacios múltiples
    name = re.sub(r'\s+', ' ', name)
    # Eliminar caracteres especiales
    name = re.sub(r'[.,;:\'"]', '', name)
    return name

def find_matching_name(target_name, name_list):
    """Encuentra el nombre más parecido en una lista de nombres."""
    if not target_name:
        return None
    
    # Intentar encontrar una coincidencia exacta
    if target_name in name_list:
        return target_name
    
    # Si no hay coincidencia exacta, buscar la más cercana
    best_match = None
    best_score = 0
    
    # Convertir nombres a conjuntos de palabras para comparar
    target_words = set(target_name.split())
    
    for name in name_list:
        name_words = set(name.split())
        
        # Calcular palabras en común
        common_words = target_words.intersection(name_words)
        
        # Calcular un puntaje de similitud
        score = len(common_words) / max(len(target_words), len(name_words))
        
        if score > best_score and score > 0.5:  # Umbral de coincidencia del 50%
            best_score = score
            best_match = name
    
    return best_match

def process_attendance(official_list_file, form_response_file, meet_file, min_required_hours=4.0, output_dir=None):
    """    
    Args:
        official_list_file: Ruta al archivo CSV con la lista oficial de estudiantes
        form_response_file: Ruta al archivo CSV con las respuestas del formulario
        meet_file: Ruta al archivo CSV con el registro de participantes del meet
        min_required_hours: Mínimo de horas requeridas para considerar asistencia
        output_dir: Directorio donde guardar los archivos de salida
    """
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    # 1. Cargar la lista oficial de estudiantes
    print(f"Cargando lista oficial de estudiantes desde {official_list_file}...")
    official_students = []
    try:
        with open(official_list_file, 'r', encoding='utf-8') as file:
            reader = csv.reader(file)
            next(reader)  # Saltar el encabezado
            for row in reader:
                if len(row) >= 2:  # Verificar que haya al menos apellidos y nombres
                    apellidos = row[0].strip() if row[0] else ""
                    nombres = row[1].strip() if row[1] else ""
                    if apellidos or nombres:  # No añadir filas vacías
                        official_students.append({
                            'apellidos': apellidos,
                            'nombres': nombres,
                            'nombre_completo': f"{apellidos} {nombres}",
                            'nombre_normalizado': normalize_name(f"{apellidos} {nombres}"),
                            'asistencia': 'F',  # Por defecto todos tienen falta
                            'en_formulario': False,
                            'horas_meet': 0.0
                        })
    except Exception as e:
        print(f"Error al cargar la lista oficial: {str(e)}")
        return
    
    print(f"Se cargaron {len(official_students)} estudiantes de la lista oficial.")
    
    # 2. Procesar el archivo de formulario
    print(f"Procesando registro del formulario desde {form_response_file}...")
    form_attendees = {}
    try:
        with open(form_response_file, 'r', encoding='utf-8') as file:
            reader = csv.reader(file)
            header = next(reader)  # Obtener encabezado
            
            # Identificar automáticamente la columna de nombres
            name_col_index = None
            for i, col in enumerate(header):
                if 'APELLIDO' in col.upper() or 'NOMBRE' in col.upper() or 'POSTULANTE' in col.upper():
                    name_col_index = i
                    break
            
            if name_col_index is None:
                name_col_index = 2  # Por defecto la columna 3 (índice 2)
            
            # Leer las respuestas del formulario
            for row in reader:
                if len(row) > name_col_index:
                    student_name = row[name_col_index].strip()
                    if student_name:
                        normalized_name = normalize_name(student_name)
                        form_attendees[normalized_name] = student_name
    except Exception as e:
        print(f"Error al procesar el formulario: {str(e)}")
        return
    
    print(f"Se encontraron {len(form_attendees)} registros en el formulario.")
    
    # 3. Procesar el archivo de participantes del Meet
    print(f"Procesando registro de participantes del Meet desde {meet_file}...")
    meet_attendance = {}
    try:
        with open(meet_file, 'r', encoding='utf-8') as file:
            reader = csv.reader(file)
            header = next(reader)  # Obtener encabezado
            
            # Identificar automáticamente las columnas
            name_cols = []
            duration_col = None
            
            for i, col in enumerate(header):
                col_lower = col.lower()
                if 'apellido' in col_lower or 'nombre' in col_lower:
                    name_cols.append(i)
                elif 'duraci' in col_lower or 'duration' in col_lower:
                    duration_col = i
            
            if not name_cols:
                name_cols = [0, 1]  # Por defecto, primeras dos columnas
            
            if duration_col is None:
                duration_col = 3  # Por defecto, cuarta columna
            
            # Leer los registros de asistencia
            for row in reader:
                if len(row) <= max(max(name_cols), duration_col):
                    continue  # Saltar filas incompletas
                
                # Construir el nombre según el formato del archivo
                name_parts = [row[i].strip() for i in name_cols if i < len(row)]
                name = " ".join([p for p in name_parts if p])  # Unir partes no vacías
                
                if not name or name.upper().startswith(('MONITOR', 'SUPERVISOR')):
                    continue  # Saltar monitores o nombres vacíos
                
                # Obtener duración
                duration = row[duration_col] if duration_col < len(row) else "0 h 0 min"
                hours = convert_duration_to_hours(duration)
                
                normalized_name = normalize_name(name)
                meet_attendance[normalized_name] = hours
    except Exception as e:
        print(f"Error al procesar el registro del Meet: {str(e)}")
        return
    
    print(f"Se encontraron {len(meet_attendance)} participantes en el Meet.")
    
    # 4. Procesar la asistencia para cada estudiante en la lista oficial
    print(f"Procesando asistencia con umbral de {min_required_hours} horas...")
    for student in official_students:
        normalized_name = student['nombre_normalizado']
        
        # Verificar registro en formulario
        form_match = find_matching_name(normalized_name, form_attendees.keys())
        if form_match:
            student['en_formulario'] = True
        
        # Verificar asistencia en Meet
        meet_match = find_matching_name(normalized_name, meet_attendance.keys())
        if meet_match:
            student['horas_meet'] = meet_attendance[meet_match]
        
        # Determinar asistencia final
        if student['en_formulario'] and student['horas_meet'] >= min_required_hours:
            student['asistencia'] = 'A'
    
    # 5. Generar el archivo de asistencia final
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    output_filename = f"asistencia_final_{timestamp}.csv"
    if output_dir:
        output_filename = os.path.join(output_dir, output_filename)
    
    try:
        with open(output_filename, 'w', newline='', encoding='utf-8') as file:
            writer = csv.writer(file)
            writer.writerow(['APELLIDOS', 'NOMBRES', 'ASISTENCIA'])
            
            for student in official_students:
                writer.writerow([student['apellidos'], student['nombres'], student['asistencia']])
    except Exception as e:
        print(f"Error al generar el archivo de asistencia: {str(e)}")
        return
    
    # 6. Generar un informe detallado
    detail_filename = f"informe_detallado_{timestamp}.csv"
    if output_dir:
        detail_filename = os.path.join(output_dir, detail_filename)
    
    try:
        with open(detail_filename, 'w', newline='', encoding='utf-8') as file:
            writer = csv.writer(file)
            writer.writerow(['APELLIDOS', 'NOMBRES', 'REGISTRO EN FORMULARIO', 'HORAS EN MEET', 'ASISTENCIA'])
            
            for student in official_students:
                writer.writerow([
                    student['apellidos'], 
                    student['nombres'], 
                    'SÍ' if student['en_formulario'] else 'NO',
                    f"{student['horas_meet']:.2f}",
                    student['asistencia']
                ])
    except Exception as e:
        print(f"Error al generar el informe detallado: {str(e)}")
    
    # 7. Generar resumen
    total_students = len(official_students)
    present_students = sum(1 for s in official_students if s['asistencia'] == 'A')
    absent_students = total_students - present_students
    
    print("\n=== RESUMEN DE ASISTENCIA ===")
    print(f"Total de estudiantes: {total_students}")
    print(f"Asistentes: {present_students} ({present_students/total_students*100:.1f}%)")
    print(f"Ausentes: {absent_students} ({absent_students/total_students*100:.1f}%)")
    print(f"Archivo de asistencia generado: {output_filename}")
    print(f"Informe detallado generado: {detail_filename}")
    
    # Mostrar estudiantes sin asistencia
    print("\nEstudiantes sin asistencia:")
    for student in official_students:
        if student['asistencia'] == 'F':
            print(f"- {student['nombre_completo']}")
            reasons = []
            if not student['en_formulario']:
                reasons.append("No registró asistencia en el formulario")
            if student['horas_meet'] < min_required_hours:
                reasons.append(f"Duración en Meet: {student['horas_meet']:.2f} horas (mínimo requerido: {min_required_hours})")
            print(f"  Motivo: {' y '.join(reasons)}")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Genera un reporte de asistencia basado en múltiples fuentes.")
    parser.add_argument('--lista', type=str, help='Ruta al archivo CSV con la lista oficial', default='asistencia prueba - Hoja 1.csv')
    parser.add_argument('--formulario', type=str, help='Ruta al archivo CSV con las respuestas del formulario', default='ASISTENCIA 19_05_2025  - Respuestas de formulario 1.csv')
    parser.add_argument('--meet', type=str, help='Ruta al archivo CSV con los registros del meet', default='19-05-2025 LUNES - Participantes.csv')
    parser.add_argument('--horas', type=float, help='Mínimo de horas requeridas para asistencia', default=4.0)
    parser.add_argument('--output', type=str, help='Directorio para archivos de salida', default=None)
   
    args = parser.parse_args()
    
    # Ejecutar procesamiento con los argumentos proporcionados
    process_attendance(
        official_list_file=args.lista,
        form_response_file=args.formulario,
        meet_file=args.meet,
        min_required_hours=args.horas,
        output_dir=args.output
    )