FROM python:3.10-slim
RUN apt-get update && apt-get install -y libreoffice && apt-get clean
WORKDIR /app
COPY . .
RUN pip install aiogram Pillow
CMD ["python", "Kirish.py"]
