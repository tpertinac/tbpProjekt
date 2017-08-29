#!/bin/bash

sudo apt-key adv --keyserver hkp://keyserver.ubuntu.com:80 --recv 0C49F3730359A14518585931BC711F9BA15703C6
echo "deb [ arch=amd64,arm64 ] http://repo.mongodb.org/apt/ubuntu xenial/mongodb-org/3.4 multiverse" | sudo tee /etc/apt/sources.list.d/mongodb-org-3.4.list
sudo apt-get update
sudo apt-get install -y mongodb-org
sudo service mongod start
sudo service mongod status
mongo
use db
PID=$!
sleep 2
kill $PID
mongoimport --db tbp --collection bitcoin --file bitcoin.json --jsonArray
mongoimport --db tbp --collection ethereum --file ethereum.json --jsonArray
mongoimport --db tbp --collection litecoin --file litecoin.json --jsonArray
mongoimport --db tbp --collection monero --file monero.json --jsonArray
mongoimport --db tbp --collection ripple --file ripple.json --jsonArray
