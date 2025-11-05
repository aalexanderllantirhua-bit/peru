#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script para crear una presentación de PowerPoint sobre
Equipos de Alineamiento Electrónico en Vehículos
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor

def crear_diapositiva_titulo(prs, titulo, subtitulo=""):
    """Crea una diapositiva de título"""
    slide_layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(slide_layout)
    
    title = slide.shapes.title
    subtitle = slide.placeholders[1]
    
    title.text = titulo
    if subtitulo:
        subtitle.text = subtitulo
    
    # Estilo del título
    title.text_frame.paragraphs[0].font.size = Pt(44)
    title.text_frame.paragraphs[0].font.bold = True
    title.text_frame.paragraphs[0].font.color.rgb = RGBColor(0, 51, 102)
    
    return slide

def crear_diapositiva_contenido(prs, titulo, contenido_lista):
    """Crea una diapositiva con título y contenido en viñetas"""
    slide_layout = prs.slide_layouts[1]
    slide = prs.slides.add_slide(slide_layout)
    
    title = slide.shapes.title
    body = slide.placeholders[1]
    
    title.text = titulo
    title.text_frame.paragraphs[0].font.size = Pt(36)
    title.text_frame.paragraphs[0].font.bold = True
    title.text_frame.paragraphs[0].font.color.rgb = RGBColor(0, 51, 102)
    
    tf = body.text_frame
    tf.clear()
    
    for item in contenido_lista:
        p = tf.add_paragraph()
        p.text = item
        p.level = 0
        p.font.size = Pt(18)
        p.space_after = Pt(12)
    
    return slide

def crear_diapositiva_dos_columnas(prs, titulo, columna_izq, columna_der):
    """Crea una diapositiva con dos columnas de contenido"""
    slide_layout = prs.slide_layouts[5]  # Layout en blanco
    slide = prs.slides.add_slide(slide_layout)
    
    # Título
    left = Inches(0.5)
    top = Inches(0.5)
    width = Inches(9)
    height = Inches(0.8)
    
    title_box = slide.shapes.add_textbox(left, top, width, height)
    title_frame = title_box.text_frame
    title_frame.text = titulo
    title_frame.paragraphs[0].font.size = Pt(36)
    title_frame.paragraphs[0].font.bold = True
    title_frame.paragraphs[0].font.color.rgb = RGBColor(0, 51, 102)
    
    # Columna izquierda
    left_col = Inches(0.5)
    top_col = Inches(1.5)
    width_col = Inches(4.25)
    height_col = Inches(5)
    
    left_box = slide.shapes.add_textbox(left_col, top_col, width_col, height_col)
    left_frame = left_box.text_frame
    left_frame.word_wrap = True
    
    for item in columna_izq:
        p = left_frame.add_paragraph()
        p.text = item
        p.font.size = Pt(16)
        p.space_after = Pt(10)
    
    # Columna derecha
    right_col = Inches(5.25)
    
    right_box = slide.shapes.add_textbox(right_col, top_col, width_col, height_col)
    right_frame = right_box.text_frame
    right_frame.word_wrap = True
    
    for item in columna_der:
        p = right_frame.add_paragraph()
        p.text = item
        p.font.size = Pt(16)
        p.space_after = Pt(10)
    
    return slide

def main():
    """Función principal para crear la presentación"""
    
    # Crear presentación
    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(7.5)
    
    # Diapositiva 1: Portada
    crear_diapositiva_titulo(
        prs,
        "Equipos de Alineamiento Electrónico en Vehículos",
        "Resumen de Uso y Aplicaciones"
    )
    
    # Diapositiva 2: ¿Qué es el Alineamiento Electrónico?
    crear_diapositiva_contenido(
        prs,
        "¿Qué es el Alineamiento Electrónico?",
        [
            "Sistema computarizado para medir y ajustar la geometría de las ruedas",
            "Utiliza sensores y cámaras de alta precisión",
            "Proporciona mediciones exactas en tiempo real",
            "Permite ajustes precisos según especificaciones del fabricante",
            "Mejora la seguridad y el rendimiento del vehículo"
        ]
    )
    
    # Diapositiva 3: Componentes del Sistema
    crear_diapositiva_contenido(
        prs,
        "Componentes del Sistema de Alineamiento",
        [
            "Consola computarizada con software especializado",
            "Cámaras de alta resolución (generalmente 4 u 8)",
            "Sensores de rueda con reflectores o targets",
            "Elevador o rampa de alineación",
            "Platos giratorios y placas deslizantes",
            "Sistema de calibración y diagnóstico"
        ]
    )
    
    # Diapositiva 4: Parámetros de Alineación
    crear_diapositiva_contenido(
        prs,
        "Parámetros Principales de Alineación",
        [
            "CAMBER: Inclinación vertical de la rueda (vista frontal)",
            "CASTER: Inclinación del eje de dirección (vista lateral)",
            "TOE (Convergencia): Ángulo horizontal de las ruedas",
            "Ángulo de empuje (Thrust Angle)",
            "SAI (Ángulo de Inclinación del Eje de Dirección)",
            "Setback (Retroceso de rueda)"
        ]
    )
    
    # Diapositiva 5: Proceso de Alineamiento
    crear_diapositiva_contenido(
        prs,
        "Proceso de Alineamiento Paso a Paso",
        [
            "1. Inspección visual previa del vehículo",
            "2. Verificación de presión de neumáticos",
            "3. Colocación del vehículo en el elevador",
            "4. Instalación de sensores en las ruedas",
            "5. Compensación de sensores (rolling compensation)",
            "6. Medición inicial de todos los parámetros",
            "7. Ajustes según especificaciones del fabricante",
            "8. Verificación final y entrega de reporte"
        ]
    )
    
    # Diapositiva 6: Ventajas vs Desventajas
    crear_diapositiva_dos_columnas(
        prs,
        "Ventajas y Consideraciones",
        [
            "✓ VENTAJAS:",
            "• Precisión milimétrica",
            "• Rapidez en el proceso",
            "• Reportes detallados",
            "• Menor margen de error",
            "• Base de datos de vehículos",
            "• Diagnóstico completo",
            "• Interfaz visual intuitiva"
        ],
        [
            "⚠ CONSIDERACIONES:",
            "• Inversión inicial elevada",
            "• Requiere capacitación",
            "• Mantenimiento periódico",
            "• Calibración regular",
            "• Espacio físico adecuado",
            "• Actualizaciones de software",
            "• Dependencia tecnológica"
        ]
    )
    
    # Diapositiva 7: Síntomas de Desalineación
    crear_diapositiva_contenido(
        prs,
        "Síntomas de Desalineación del Vehículo",
        [
            "Desgaste irregular o prematuro de neumáticos",
            "Vehículo se desvía hacia un lado",
            "Volante descentrado al conducir en línea recta",
            "Vibración en el volante",
            "Manejo inestable o impreciso",
            "Mayor consumo de combustible",
            "Ruidos anormales en la suspensión"
        ]
    )
    
    # Diapositiva 8: Tipos de Equipos
    crear_diapositiva_contenido(
        prs,
        "Tipos de Equipos de Alineamiento",
        [
            "Sistemas 3D: Utilizan cámaras y targets reflectivos",
            "Sistemas CCD: Sensores de carga acoplada",
            "Sistemas láser: Tecnología de rayos láser",
            "Sistemas de 4 cámaras: Configuración estándar",
            "Sistemas de 8 cámaras: Mayor precisión y rapidez",
            "Sistemas móviles: Portátiles para servicio en sitio"
        ]
    )
    
    # Diapositiva 9: Mantenimiento del Equipo
    crear_diapositiva_contenido(
        prs,
        "Mantenimiento del Equipo",
        [
            "Calibración mensual o según uso",
            "Limpieza de cámaras y sensores",
            "Verificación de platos y placas deslizantes",
            "Actualización de software y base de datos",
            "Revisión de conexiones y cables",
            "Protección contra polvo y humedad",
            "Capacitación continua del personal"
        ]
    )
    
    # Diapositiva 10: Beneficios para el Cliente
    crear_diapositiva_contenido(
        prs,
        "Beneficios para el Cliente",
        [
            "Mayor vida útil de los neumáticos (hasta 30% más)",
            "Mejor economía de combustible",
            "Conducción más segura y estable",
            "Menor desgaste de componentes de suspensión",
            "Reporte técnico detallado y profesional",
            "Garantía de trabajo realizado",
            "Prevención de problemas futuros"
        ]
    )
    
    # Diapositiva 11: Marcas Principales
    crear_diapositiva_contenido(
        prs,
        "Principales Fabricantes de Equipos",
        [
            "Hunter Engineering - Líder mundial en alineación",
            "John Bean - Tecnología italiana de precisión",
            "Hofmann - Equipos alemanes de alta calidad",
            "Ravaglioli - Soluciones profesionales",
            "Corghi - Innovación y diseño italiano",
            "Launch - Tecnología asiática accesible",
            "Beissbarth - Precisión alemana"
        ]
    )
    
    # Diapositiva 12: Conclusiones
    crear_diapositiva_contenido(
        prs,
        "Conclusiones",
        [
            "El alineamiento electrónico es esencial en talleres modernos",
            "Inversión que se recupera con el volumen de trabajo",
            "Mejora la satisfacción y confianza del cliente",
            "Tecnología en constante evolución",
            "Requiere personal capacitado y comprometido",
            "Fundamental para la seguridad vial",
            "Diferenciador competitivo en el mercado"
        ]
    )
    
    # Diapositiva 13: Cierre
    crear_diapositiva_titulo(
        prs,
        "¡Gracias!",
        "Preguntas y Comentarios"
    )
    
    # Guardar presentación
    nombre_archivo = "Presentacion_Alineamiento_Electronico.pptx"
    prs.save(nombre_archivo)
    print(f"✓ Presentación creada exitosamente: {nombre_archivo}")
    print(f"✓ Total de diapositivas: {len(prs.slides)}")
    
    return nombre_archivo

if __name__ == "__main__":
    main()
