# ZoraEC · Simulador Financiero 360

Versión del producto: **1.0.0**  
Diseño y desarrollo: **Freddy Beltrán A.**  
Marca: **ZoraEC**

## Historial

### 1.0.0 — 2026-08-27

Primera versión funcional designada por el autor.

- Simulación de préstamos hipotecarios y comerciales con sistemas francés y alemán.
- Proyección de ahorro compuesto y aportes mensuales.
- Resúmenes de resultados, capacidad de pago y cumplimiento de metas.
- Vista previa única con formato monetario.
- Indicadores distribuidos en dos filas para mejorar la lectura.
- Detalles y gráficos contraíbles, con indicadores +/−.
- Exportación Excel con marca, versión y crédito de desarrollo.
- Identidad y versión centralizadas en constantes del archivo Python.

## Regla para las siguientes versiones

- 1.0.1: correcciones compatibles, por ejemplo un ajuste visual.
- 1.1.0: nuevas funciones compatibles, por ejemplo una comparación adicional.
- 2.0.0: cambios incompatibles en entradas, salidas o comportamiento documentado.

Referencia: https://semver.org/lang/es/

Actualizar VERSION en el código y agregar una entrada al historial por cada entrega.
No sobrescribir una versión ya publicada: asignar un número nuevo.

## Control de cambios recomendado

Usar Git en una carpeta exclusiva del simulador, con un archivo estable como app.py,
un README y este historial. Evitar nombres como final, final2 o corregido2.
No inicializar un repositorio en la carpeta zora completa: contiene otros proyectos.
Este documento y la constante VERSION no sustituyen el historial de Git.
No se ha creado un repositorio, commit, etiqueta ni publicación remota en esta tarea.

Después de verificar la aplicación, registrar el código en un commit y marcarlo
con la etiqueta v1.0.0. Conservar las versiones exactas de las dependencias del
entorno donde se haya validado, sin inventar versiones ni fijar paquetes no probados.

## Comprobación previa a publicación

- Ejecutar los tres simuladores y verificar los resultados esperados.
- Revisar importes grandes y visualización en móvil y escritorio.
- Abrir y cerrar cada sección; comprobar los signos +/−.
- Descargar y abrir los reportes de préstamos y ahorro en Excel.
- Revisar fórmulas, créditos y versión del reporte.

La validación visual en Streamlit sigue pendiente en el entorno del usuario.
