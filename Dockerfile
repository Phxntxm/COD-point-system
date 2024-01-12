FROM python:3.11

WORKDIR /app

COPY Pipfile Pipfile.lock /app/

RUN pip install --upgrade pip
RUN pip install pipenv
RUN pipenv install --system --deploy

COPY main.py /app/

COPY src /app/src

CMD ["python", "-u", "main.py"]