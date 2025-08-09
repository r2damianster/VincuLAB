# Sistema Unificado de Gestión de Reportes - Educación Especial

## 📋 Descripción

El **Sistema Unificado de Gestión de Reportes de Educación Especial** es una aplicación Python con interfaz gráfica (Tkinter) diseñada para automatizar la generación de reportes consolidados, consultas de estudiantes, análisis de beneficiarios y creación de oficios institucionales para proyectos de vinculación con la sociedad en el área de educación especial.

## ✨ Características Principales

- **🔄 Generación Automática de Reportes**: Crea reportes consolidados con datos de múltiples fuentes
- **👥 Consulta de Estudiantes**: Sistema de búsqueda avanzada con filtros por período
- **📊 Análisis de Beneficiarios**: Genera gráficos y análisis estadísticos
- **📄 Oficios Institucionales**: Crea documentos oficiales personalizados automáticamente
- **🔐 Sistema de Autenticación**: Control de acceso mediante contraseña
- **📱 Interfaz Intuitiva**: GUI amigable desarrollada en Tkinter
- **🌐 Integración con Google Sheets**: Conexión directa con hojas de cálculo en línea

## 🛠️ Tecnologías Utilizadas

- **Python 3.11+**
- **Tkinter** - Interfaz gráfica
- **Pandas** - Manipulación de datos
- **Matplotlib/Seaborn** - Visualización de datos
- **python-docx** - Generación de documentos Word
- **WordCloud** - Nubes de palabras
- **thefuzz** - Búsqueda difusa de texto
- **openpyxl** - Manejo de archivos Excel

## 📦 Instalación

### Requisitos Previos

- Python 3.11 o superior
- pip (gestor de paquetes de Python)

### Instalación de Dependencias

```bash
pip install pandas matplotlib seaborn python-docx wordcloud thefuzz openpyxl pillow
```

### Archivos Requeridos

1. **Código principal**: `sistema_reportes.py`
2. **Plantilla de oficio**: `Formato Oficio - Editable.docx`

## 🚀 Uso Rápido

1. **Ejecutar la aplicación**:
   ```bash
   python sistema_reportes.py
   ```

2. **Configurar período**: Ingresa el período en formato `YYYY-Q` (ejemplo: `2025-1`)

3. **Seleccionar carpeta de destino**: Elige dónde guardar los archivos generados

4. **Usar las funciones disponibles**:
   - Generar Reporte Unificado Consolidado
   - Consultar Estudiante
   - Generar Análisis de Beneficiarios
   - Generar Oficios Institucionales

## 📊 Fuentes de Datos

El sistema se conecta automáticamente a las siguientes fuentes de datos:

### 1. **Base de Instituciones**
- **URL**: Google Sheets principal de instituciones
- **Columnas clave**: 
  - `Período` (timestamp convertido a YYYY-Q)
  - `Nombre Completo de la Institución`
  - `Nombre Corto de la Institución`
  - `Rector(a) o Autoridad`
  - `Supervisor`
  - `Proyecto`

### 2. **Base de Beneficiarios**
- **URL**: Google Sheets de beneficiarios
- **Columnas clave**:
  - `Período de registro` (formato YYYY-Q)
  - `Centro de Educación`
  - `Qué voy a reportar`
  - `Número total de participantes en la sensibilización`
  - `Número de padres y cuidadores Capacitación`

### 3. **Base de Ubicaciones**
- **URL**: Google Sheets de ubicaciones geográficas
- **Columnas clave**:
  - `INSTITUCIÓN`
  - `LATITUD`
  - `LONGITUD`

### 4. **Base de Encuestas**
- **URL**: Google Sheets de encuestas de satisfacción
- **Uso**: Análisis de beneficiarios y retroalimentación

### 5. **Base de Contraseñas**
- **URL**: Google Sheets con credenciales de acceso
- **Uso**: Autenticación para funciones sensibles



## 🏗️ Estructura del Código

### Clase Principal: `UnifiedReportApp`

```python
class UnifiedReportApp:
    def __init__(self, root)                    # Inicialización de la interfaz
    def create_widgets(self)                    # Creación de elementos GUI
    def load_data(self)                         # Carga de datos desde fuentes
    def convert_timestamp_to_period(self)       # Conversión de fechas a formato YYYY-Q
    def generate_unified_consolidated_report(self)  # Generación de reportes consolidados
    def start_student_consultation(self)        # Inicio de consulta de estudiantes
    def generate_beneficiary_analysis(self)     # Análisis de beneficiarios
    def generate_official_documents(self)       # Generación de oficios
```

### Funciones Clave

#### 🔄 **Conversión de Períodos**
```python
def convert_timestamp_to_period(self, timestamp_value):
    """Convierte timestamps a formato YYYY-Q"""
    # Convierte fechas como "2025-03-15" a "2025-1"
    # Maneja diferentes formatos de entrada
```

#### 📊 **Generación de Reportes**
```python
def generate_unified_consolidated_report(self):
    """Genera reporte consolidado con múltiples hojas"""
    # Hoja 1: Beneficiarios (resumen por centro)
    # Hoja 2: Actividades (actividades por centro)  
    # Hoja 3: Centros_Educacion (información detallada filtrada por período)
```

#### 🔍 **Búsqueda de Estudiantes**
```python
def search_student(self, df_estudiantes, nombre_busqueda):
    """Búsqueda exacta y difusa de estudiantes"""
    # Búsqueda exacta por coincidencia de texto
    # Búsqueda difusa usando algoritmo de similitud
    # Filtrado opcional por período
```

#### 📄 **Generación de Oficios**
```python
def create_official_document(self, ...):
    """Crea documentos oficiales personalizados"""
    # Reemplazo robusto de marcadores en documentos Word
    # Mapeo automático de datos desde múltiples fuentes
    # Validación y limpieza de datos
```

## 🎯 Funcionalidades Detalladas

### 1. **📈 Reporte Unificado Consolidado**

**Propósito**: Genera un archivo Excel con tres hojas que consolidan toda la información del período.

**Proceso**:
1. Filtra datos de beneficiarios por período (`Período de registro`)
2. Filtra datos de instituciones por período (`Período` convertido a YYYY-Q)
3. Calcula métricas por categorías:
   - Atenciones Individuales
   - Asesorías
   - Evaluaciones Psicopedagógicas
   - DIAC (Documentos de Adaptación Curricular)
   - Capacitaciones a Funcionarios
   - Sensibilizaciones
   - Capacitaciones a Padres

**Salida**:
- `Reporte_Unificado_Consolidado_YYYY-Q.xlsx`
  - **Hoja "Beneficiarios"**: Resumen de beneficiarios por centro
  - **Hoja "Actividades"**: Resumen de actividades por centro
  - **Hoja "Centros_Educacion"**: Información detallada de instituciones (FILTRADA por período)

### 2. **👥 Consulta de Estudiantes**

**Propósito**: Búsqueda avanzada de información de estudiantes con autenticación.

**Características**:
- **Autenticación**: Requiere contraseña para acceso
- **Búsqueda exacta**: Por coincidencia de texto
- **Búsqueda difusa**: Usando algoritmo de similitud (score > 60)
- **Filtro por período**: Opcional
- **Información completa**: Datos académicos, contacto, ubicación geográfica
- **Enlaces a Google Maps**: Si hay coordenadas disponibles

**Proceso**:
1. Verificación de contraseña desde Google Sheets
2. Carga de datos de estudiantes con conversión de períodos
3. Merge con datos de ubicación geográfica
4. Búsqueda con múltiples algoritmos
5. Presentación de resultados en ventana con scroll

### 3. **📊 Análisis de Beneficiarios**

**Propósito**: Genera análisis estadístico y visual de encuestas de satisfacción.

**Componentes**:
- **Gráficos estadísticos**:
  - Distribución por ciudad
  - Distribución por institución
- **Nubes de palabras**:
  - Aspectos positivos
  - Aspectos a mejorar
- **Documento Word**: Compilación de todos los análisis

**Salida**:
- Carpeta `Beneficiarios/`
  - `Encuesta_Beneficiarios_YYYY-Q.xlsx`
  - `ciudad_distribution_YYYY-Q.jpg`
  - `institucion_distribution_YYYY-Q.jpg`
  - `wordcloud_positive_YYYY-Q.png`
  - `wordcloud_improvement_YYYY-Q.png`
  - `Beneficiarios_analisis_YYYY-Q.docx`

### 4. **📄 Oficios Institucionales**

**Propósito**: Genera documentos oficiales personalizados para cada institución.

**Características**:
- **Filtrado por período**: Solo instituciones del período solicitado
- **Reemplazo robusto**: Maneja documentos Word complejos
- **Datos automáticos**: Obtiene información desde múltiples fuentes
- **Validación**: Verifica existencia de datos antes de generar

**Marcadores reemplazados**:
```
[Fecha] → Fecha actual
[Título del representante de la institución] → Desde Google Sheets
[Nombre del Representante de la institucion] → Desde columna "Rector(a) o Autoridad"
[Cargo del representante] → Desde Google Sheets
[Nombre Completo de la Institución] → Desde Google Sheets
[Supervisor del Proyecto] → Desde columna "Supervisor"
[Proyecto] → Desde Google Sheets

[Número de Capacitaciones Funcionarios] → Desde reporte consolidado
[Número de Funcionarios Capacitados] → Calculado automáticamente
[Número de Capacitaciones Padres] → Desde reporte consolidado
[Número de Padres Capacitados] → Desde reporte consolidado
[Número de Sensibilizaciones] → Desde reporte consolidado
[Número de Personas Sensibilizadas] → Desde reporte consolidado
[Número de Asesorías] → Desde reporte consolidado
[Atenciones Individuales] → Desde reporte consolidado
[Estudiantes DIAC] → Desde reporte consolidado
[Evaluaciones Psicopedagógicas] → Desde reporte consolidado
```

**Proceso**:
1. Verificación de plantilla `Formato Oficio - Editable.docx`
2. Verificación/generación de reporte consolidado
3. Filtrado de instituciones por período
4. Mapeo de datos desde múltiples fuentes
5. Reemplazo robusto de marcadores
6. Generación de archivos individuales

**Salida**:
- Carpeta `Oficios_Instituciones/`
  - `Oficio_[Institución]_No_[Número]_YYYY-Q.docx` (uno por institución)


## ⚙️ Configuración

### URLs de Fuentes de Datos

Las URLs están configuradas en la clase `UnifiedReportApp.__init__()`:

```python
# URLs de datos (configurables)
self.url_datos = "https://docs.google.com/spreadsheets/d/[ID]/export?format=xlsx"
self.url_beneficiarios = "https://docs.google.com/spreadsheets/d/[ID]/export?format=csv"
self.url_ubicacion = "https://docs.google.com/spreadsheets/d/[ID]/export?format=xlsx"
self.url_encuesta = "https://docs.google.com/spreadsheets/d/[ID]/export?format=csv"
self.url_password = "https://docs.google.com/spreadsheets/d/[ID]/export?format=csv"
```

### Formato de Períodos

El sistema utiliza el formato **YYYY-Q** donde:
- **YYYY**: Año de 4 dígitos
- **Q**: Trimestre (1, 2, 3, 4)
- **Ejemplos**: `2025-1`, `2024-2`, `2023-4`

### Conversión Automática de Períodos

```python
# Conversión de timestamps a formato YYYY-Q
Enero-Marzo → Q1    (2025-01-15 → 2025-1)
Abril-Junio → Q2    (2025-05-20 → 2025-2)
Julio-Septiembre → Q3    (2025-08-10 → 2025-3)
Octubre-Diciembre → Q4   (2025-11-30 → 2025-4)
```

## 🔧 Mapeo de Columnas

### Diferencias entre Fuentes de Datos

El sistema maneja automáticamente las diferencias en nombres de columnas:

| Concepto | Base Instituciones | Base Beneficiarios | Base Ubicaciones |
|----------|-------------------|-------------------|------------------|
| **Período** | `Período` (timestamp) | `Período de registro` (YYYY-Q) | N/A |
| **Institución** | `Nombre Corto de la Institución` | `Centro de Educación` | `INSTITUCIÓN` |
| **Representante** | `Rector(a) o Autoridad` | N/A | N/A |
| **Supervisor** | `Supervisor` | N/A | N/A |
| **Coordenadas** | N/A | N/A | `LATITUD`, `LONGITUD` |

### Categorías de Beneficiarios

```python
categorias = {
    'Atenciones Individuales': 'Estudiante atendido Individualmente',
    'Asesorías': 'Asesorías a funcionarios',
    'Evaluaciones Psicopedagógicas': 'Beneficiarios de evaluación psicopedagógica',
    'DIAC': 'Beneficiarios DIAC o plan de intervención',
    'Capacitaciones Funcionarios': 'Capacitación a funcionario(s) (Docentes u otros)',
    'Sensibilizaciones': 'Sensibilización',
    'Capacitaciones Padres': 'Capacitación a padres de familia'
}
```

## 🐛 Solución de Problemas

### Problemas Comunes y Soluciones

#### 1. **Error: "No se encontró la columna 'Período de registro'"**

**Causa**: La columna no existe o tiene un nombre diferente en la base de beneficiarios.

**Solución**:
```python
# Verificar nombres exactos de columnas
print(f"Columnas disponibles: {list(df_beneficiarios.columns)}")
```

#### 2. **Error: "No se encontraron instituciones para el período YYYY-Q"**

**Causa**: No hay instituciones registradas para ese período o el formato es incorrecto.

**Solución**:
- Verificar formato del período: `2025-1` (no `2025-01`)
- Verificar que existan datos para ese período
- Revisar la conversión de timestamps

#### 3. **Oficios con marcadores sin reemplazar**

**Causa**: Los marcadores en la plantilla no coinciden exactamente con los definidos en el código.

**Solución**:
- Verificar que la plantilla `Formato Oficio - Editable.docx` existe
- Verificar que los marcadores están escritos exactamente como: `[Marcador]`
- Revisar la información de depuración en consola

#### 4. **Error de conexión a Google Sheets**

**Causa**: Problemas de conectividad o URLs incorrectas.

**Solución**:
- Verificar conexión a internet
- Verificar que las URLs de Google Sheets son públicas
- Verificar formato de exportación (xlsx vs csv)

#### 5. **Búsqueda de estudiantes no encuentra resultados**

**Causa**: Nombre mal escrito o estudiante no existe en el período.

**Solución**:
- Usar búsqueda difusa (el sistema la hace automáticamente)
- Verificar filtro de período
- Revisar nombres exactos en la base de datos

### Información de Depuración

El sistema proporciona información detallada en consola:

```python
# Ejemplo de salida de depuración
Columnas disponibles en beneficiarios: ['Período de registro', 'Centro de Educación', ...]
Buscando período: 2025-1
Registros encontrados para 2025-1: 150
Períodos únicos en instituciones: ['2023-1', '2023-2', '2024-1', '2024-2', '2025-1']
Instituciones encontradas para 2025-1: 15
✅ Centros después del filtrado por período: 15
🏢 Generando oficio para: Unidad Educativa Ejemplo
👤 Nombre del representante obtenido: 'Dr. Juan Pérez'
👨‍🏫 Supervisor obtenido desde BD: 'Mg. María González'
🔄 Reemplazos definidos:
  [Nombre del Representante de la institucion] → 'Dr. Juan Pérez'
  [Supervisor del Proyecto] → 'Mg. María González'
✅ Documento generado exitosamente
```

## 📝 Mejores Prácticas

### 1. **Preparación de Datos**

- **Verificar períodos**: Asegurar que los datos estén en el período correcto
- **Validar nombres**: Los nombres de instituciones deben ser consistentes entre fuentes
- **Revisar plantilla**: La plantilla de oficio debe tener todos los marcadores necesarios

### 2. **Uso del Sistema**

- **Generar reporte primero**: Siempre generar el reporte consolidado antes de los oficios
- **Verificar carpeta destino**: Asegurar que la carpeta de destino sea accesible
- **Revisar información de depuración**: Usar la consola para verificar el procesamiento

### 3. **Mantenimiento**

- **Actualizar URLs**: Si cambian las URLs de Google Sheets, actualizar en el código
- **Verificar columnas**: Si cambian nombres de columnas, actualizar el mapeo
- **Backup de plantillas**: Mantener copias de seguridad de las plantillas de documentos

## 🔒 Seguridad

### Sistema de Autenticación

- **Consulta de estudiantes**: Requiere contraseña almacenada en Google Sheets
- **Intentos limitados**: Máximo 3 intentos de contraseña
- **Datos sensibles**: La información de estudiantes está protegida

### Privacidad de Datos

- **Conexiones HTTPS**: Todas las conexiones a Google Sheets usan HTTPS
- **Datos locales**: Los archivos se guardan localmente en la carpeta seleccionada
- **Sin almacenamiento permanente**: Los datos se cargan dinámicamente

## 📊 Métricas y Estadísticas

### Datos Procesados Típicos

- **Instituciones**: 15-50 por período
- **Beneficiarios**: 100-500 registros por período
- **Oficios generados**: 1 por institución activa
- **Tiempo de procesamiento**: 30-60 segundos para reporte completo

### Formatos de Salida

- **Excel**: `.xlsx` (reportes consolidados, encuestas)
- **Word**: `.docx` (oficios, análisis)
- **Imágenes**: `.jpg`, `.png` (gráficos, nubes de palabras)

## 🤝 Contribución

### Estructura para Nuevas Funcionalidades

1. **Agregar función en la clase**: `UnifiedReportApp`
2. **Crear botón en la interfaz**: `create_widgets()`
3. **Implementar lógica de datos**: Seguir patrones existentes
4. **Agregar validación**: Verificar datos antes de procesar
5. **Incluir información de depuración**: Para facilitar el mantenimiento

### Convenciones de Código

- **Nombres de funciones**: `snake_case`
- **Nombres de variables**: `snake_case`
- **Comentarios**: En español para funciones principales
- **Información de depuración**: Usar emojis para facilitar lectura (`🏢`, `📊`, `✅`, `❌`)

## 📞 Soporte

### Información de Contacto

Para soporte técnico o consultas sobre el sistema, contactar al equipo de desarrollo del proyecto de vinculación de la Carrera de Educación Especial.

### Logs y Depuración

El sistema genera información detallada en la consola. Para reportar problemas, incluir:

1. **Mensaje de error completo**
2. **Período utilizado**
3. **Información de depuración de la consola**
4. **Pasos para reproducir el problema**

---

## 📄 Licencia

Este sistema fue desarrollado para la **Universidad Laica Eloy Alfaro de Manabí - Carrera de Educación Especial** como parte del proyecto de vinculación con la sociedad "Espacios de Apoyo Pedagógico Inclusivo".

---

**Versión**: 2.0  
**Última actualización**: Enero 2025  
**Desarrollado por**: Equipo de Vinculación - Educación Especial ULEAM

