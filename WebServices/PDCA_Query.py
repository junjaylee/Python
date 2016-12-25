#!/usr/bin/env python
#http://localhost:1234/ProductID=Kirin&SerialNumber=C7CRW01UHGJR&StationID=DEVELOPMENT20&StartTime=2016-07-08 00:00:00&EndTime=2016-07-09 00:00:00
#http://localhost:1234/StationID=DEVELOPMENT20&EndTime=2016-07-09+00%3A00%3A00&StartTime=2016-07-08+00%3A00%3A00&ProductID=Kirin
import sys
import pymongo
import plistlib
import logging
import json

PORT_NUMBER = 1234
if __name__ == "__main__":
    if (sys.version_info > (3, 0)):
        from http.server import BaseHTTPRequestHandler,HTTPServer
        from socketserver import ThreadingMixIn
        import urllib.parse
    else:
        from BaseHTTPServer import BaseHTTPRequestHandler,HTTPServer
        from SocketServer import ThreadingMixIn
        import urlparse

class myHandler(BaseHTTPRequestHandler):
    def _set_headers(self,pCode):
        self.send_response(pCode)
        self.send_header('Content-type', 'application/json')
        self.end_headers()

    def ParseCommandList(self,pCommand,pDcit):
        dictTime = dict()
        pKeys = ['SerialNumber','ProductID','StationID']
        if (sys.version_info > (3, 0)):
            DecodeUrl = urllib.parse.parse_qsl(pCommand)
        else:
            DecodeUrl = urlparse.parse_qsl(pCommand)
        for key in DecodeUrl:
            if (len(key) == 2):
                if key[0] == 'StartTime':
                    dictTime['$gte'] = key[1]                   
                elif key[0] == 'EndTime':
                    dictTime['$lt'] = key[1]
                elif key[0] in pKeys:
                    pDcit.update({key[0]:key[1]})
        return dictTime

    def do_GET(self):
        Tool_Ver = 1.0
        Reporting = dict()
        logging.basicConfig(format='%(asctime)-15s %(levelname)s => %(message)s',filename='PDCA_Query.log',level=logging.DEBUG)
        logging.info("Tool Ver. = %s" % Tool_Ver)
        PlistName = "PDCA_Query.plist"
        pl = plistlib.readPlist(PlistName)
        Db_Url = pl["MongoDB_URL"]
        collect = pymongo.uri_parser.parse_uri(Db_Url)
        logging.info("Query Paramter : {0}".format(self.path))
        try:
            mng_client = pymongo.MongoClient(collect['nodelist'][0][0],collect['nodelist'][0][1])
            mng_client.admin.authenticate(collect['username'],collect['password'])
            mng_db = mng_client['ndcweb']
            cmd_list = self.path[1:]
            dictInfo = dict()
            dictTime = self.ParseCommandList(cmd_list,dictInfo)
            if (len(dictInfo) >= 2):
                TblName = '{0}_{1}_rawdata'.format(dictInfo["ProductID"],dictInfo["StationID"])
                db_cm = mng_db[TblName]
                print(db_cm)
                NameList = '{0}_{1}_tinamean'.format(dictInfo["ProductID"],dictInfo["StationID"])
                QryCmdDict = dict(dictInfo)
                QryCmdDict.pop("StationID")
                QryCmdDict.pop("ProductID")
                if (len(dictTime) > 0):
                    QryCmdDict['StartTime'] = dictTime 
                logging.info("MongoDB Paramter : {0}".format(QryCmdDict))
                my_cursor = db_cm.find(QryCmdDict)
                Reporting = list()
                logging.info("result : {0}".format(my_cursor.count()))
                for document in my_cursor:
                    dictDoc = dict()
                    db_name = mng_db[NameList]
                    for szKey in document.keys():
                        if szKey == '_id':
                            continue
                        cursor = db_name.find({'redefinename':szKey})
                        dictDoc.update({cursor[0]['originalname']:document[szKey]})
                    Reporting.append(dictDoc)
            mng_client.close()
            logging.info("Finished")
            self._set_headers(200)
            json_str = json.dumps(Reporting)
            # Send message back to client
            if (sys.version_info > (3, 0)):
                self.wfile.write(bytes(json_str, "utf8"))
            else:
                self.wfile.write(json_str)
        except pymongo.errors.ConnectionFailure as e:
            logging.error("Connection Failure %s" % e)
            self._set_headers(999)
        except Exception as e:
            logging.error("%s" % e)
            self._set_headers(998)

class ThreadedHTTPServer(ThreadingMixIn, HTTPServer):
    """Handle requests in a separate thread."""

if __name__ == "__main__":
    try:
        #url_params = urllib.urlencode({'ProductID':'Kirin','StationID':'DEVELOPMENT20','StartTime':'2016-07-08 00:00:00','EndTime':'2016-07-09 00:00:00'})
        #print(url_params)
        #Create a web server and define the handler to manage the
        #incoming request
        server = ThreadedHTTPServer(('localhost', PORT_NUMBER), myHandler)
        print('Started httpserver on port {0}'.format(PORT_NUMBER))
        #Wait forever for incoming htto requests
        server.serve_forever()
    except KeyboardInterrupt:
        print('^C received, shutting down the web server')
        server.socket.close()
