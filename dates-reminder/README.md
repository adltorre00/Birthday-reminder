# Contenido del archivo: /dates-reminder/dates-reminder/README.md

# Proyecto de Recordatorio de Fechas

Este proyecto es un recordatorio de fechas que permite gestionar cumpleaños y eventos importantes. El script principal carga datos desde un archivo Excel, verifica los eventos y cumpleaños próximos, y envía recordatorios por correo electrónico.

## Estructura del Proyecto

- **DatesReminder.py**: Contiene la lógica principal para cargar datos, verificar eventos y cumpleaños, y enviar correos electrónicos.
- **requirements.txt**: Lista las dependencias necesarias para ejecutar el script de Python.
- **.github/workflows/schedule.yml**: Configura un flujo de trabajo de GitHub Actions que se ejecuta automáticamente cada dos días.

## Requisitos

Para ejecutar este proyecto, asegúrate de tener instaladas las siguientes bibliotecas:

- pandas
- smtplib

Puedes instalar las dependencias ejecutando:

```
pip install -r requirements.txt
```

## Configuración

1. Asegúrate de tener un archivo `Dates.xlsx` en el mismo directorio que `DatesReminder.py`, con las hojas "Birthdates" y "Events".
2. Configura las credenciales de correo electrónico en el script para enviar los recordatorios.

## Ejecución

El script se ejecuta automáticamente cada dos días gracias a la configuración en el archivo `schedule.yml`. Puedes ejecutarlo manualmente con:

```
python DatesReminder.py
```

## Contribuciones

Las contribuciones son bienvenidas. Si deseas mejorar este proyecto, por favor abre un issue o un pull request.