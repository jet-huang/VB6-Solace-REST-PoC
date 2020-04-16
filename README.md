# VB6-Solace-REST-PoC
This shows how we can communicate old-school VB6 (Visual Basic 6) with the latest technology.

## Before you start
1. Obtain the connection to Solace PS+ Broker. You may:  
	- Follow [these instructions](https://cloud.solace.com/learn/group_getting_started/ggs_signup.html) to quickly spin up a cloud-based Solace messaging service for your applications.
	- Follow [these instructions](https://docs.solace.com/Solace-SW-Broker-Set-Up/Setting-Up-SW-Brokers.htm) to start the Solace VMR in leading Clouds, Container Platforms or Hypervisors. The tutorials outline where to download and how to install the Solace VMR.
	- If your company has Solace message routers deployed, contact your middleware team to obtain the host name or IP address of a Solace message router to test against, a username and password to access it, and a VPN in which you can produce and consume messages.

2. Check our nice REST messaging samples [here](https://solace.com/samples/solace-samples-rest-messaging/).

3. If you want to learn more on Solace REST messaging, check [here](https://docs.solace.com/Open-APIs-Protocols/REST-get-start.htm).

## Quick Start
1. Clone this repository.

2. Configuring REST Delivery Points on your broker:  
	- Connect to broker CLI.
	- Open Create_REST.cli in any text editor.
	- Copy all content in that file and paste to your CLI prompt.
	- Check your RDPs with CLI or Web UI. There should be 5 queues and 5 RDPs.

3. You can start the pre-compiled binary "

3. _(OPTIONAL)_ Create databases with Docker:
- MariaDB:  
```shell
docker run --name vb6-mariadb -p 3306:3306 \
-e MYSQL_ROOT_PASSWORD=Solace1234 -e MYSQL_DATABASE=vb6-db \
-e MYSQL_USER=vb6 -e MYSQL_PASSWORD=Solace1234 \
-d mariadb:10.5 \
--character-set-server=utf8mb4 --collation-server=utf8mb4_unicode_ci \
--bind-address=0.0.0.0
```
	
- Microsoft SQL Server (2017):  

```shell
docker run --name vb6-mssqldb -p 1433:1433 \
-e "ACCEPT_EULA=Y" -e "SA_PASSWORD=Solace1234" \
-e "MSSQL_PID=Express" \
-d mcr.microsoft.com/mssql/server:2017-latest
```

TO BE CONTINUED, PLEASE WAIT...
