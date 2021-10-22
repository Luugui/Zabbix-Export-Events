# Zabbix Export Events

![](https://img.shields.io/badge/Zabbix%20API-4-green)

Exporta os alertas gerados no Zabbix baseado no periodo e grupos selecionados.

## Requisitos
É necessário ter os pacotes pyzabbix, openpyxl, argparse, tqdm

`pip install -r requirements.txt`


------------

## Exemplo
É obrigatório passar o servidor, login e senha do qual os dados serão extraídos.

`python .\get_events.py -s http://localhost/zabbix -u Admin -p zabbix`

O parâmetro -n [Name] é o nome que será salvo o relatório junto com a data atual

`python .\get_events.py -s http://localhost/zabbix -u Admin -p zabbix -n CLIENT`

O parâmetro -g [Group] serve para especificar o grupo do qual os alertas serão extraidos. Caso tenha subgrupos ele irá incluir eles na listagem, na ausência desse parâmetro ele irá extrair de todos os grupos que tiver permissão de leitura.

`python .\get_events.py -s http://localhost/zabbix -u Admin -p zabbix -n CLIENT -g "NOME GRUPO"`

Existem duas formas de configurar o periodo de extração. Podemos especificar uma data através dos parametros --data-inicio e --data-fim ou simplesmente dizendo das ultimas X horas através do comando --last

`python .\get_events.py --server http://localhost/zabbix --user Admin --password zabbix --name CLIENT --group "NOME GRUPO" --last 1`

No caso do uso do --data-inicio / --data-fim é possivel especificar data e hora da extração no formato "DD/MM/AA HH:MM:SS" sendo que a especificação do horario não é obrigatório, caso não seja passado irá utilizar 00:00:00.

`python .\get_events.py --server http://localhost/zabbix --user Admin --password zabbix --name CLIENT --group "NOME GRUPO" --data-inicio 01/03/2020 --data-fim 01/04/2020`

Com a flag --ack você pode selecionar apenas os eventos que receberam acknowleged.

`python .\get_events.py --server http://localhost/zabbix --user Admin --password zabbix --name CLIENT --group "NOME GRUPO" --last 1 --email fulano@company.com --ack`

------------

#### Configuração do envio de email

O Script também pode realizar o envio do relatório gerado para um endereço de email.

Edite as linhas 340 e 341 com a conta para envio e na linha 342 com o endereço do servidor smtp.

	email_user = "account@company.com.br"  # Account used for send e-mail
    password = "password"  # Account passowrd used to send e-mail
    smtp_adrress = "smtp.company.com.br"  # SMTP Server
    email_send = args["email"]  # E-mail to receve e-mail

Após configurado para envio do relatório basta utilizar o parametro --email para especificar qual endereço irá receber o relatório.

`python .\get_events.py --server http://localhost/zabbix --user Admin --password zabbix --name CLIENT --group "NOME GRUPO" --last 1 --email fulano@company.com`


