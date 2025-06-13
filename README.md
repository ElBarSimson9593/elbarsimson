# Document Generator

This repository provides a small FastAPI application to generate corporate reports and PowerPoint presentations.

## Requisitos

- Python 3.11 (incluye FastAPI en este entorno)

## Uso

Instale las dependencias y ejecute el servidor:

```bash
pip install openai
python server.py
```

Abra `http://localhost:8000` en su navegador para ingresar el título y contenido del informe o la presentación.

- El botón **Crear Informe** genera un archivo HTML de más de 30 páginas que puede imprimir como PDF.
- El botón **Crear PPT** descarga un archivo `presentation.pptx` sencillo con una diapositiva que contiene el título y el contenido.

## Uso con LocalAI

Si dispone de una instancia de [LocalAI](https://github.com/go-skynet/LocalAI), defina la variable de entorno `OPENAI_BASE_URL` apuntando a su servidor (por ejemplo `http://localhost:8080/v1`). El programa utilizará esa URL para generar el texto de los informes.
