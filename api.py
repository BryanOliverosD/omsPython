from flask import Flask
from flask_restful import Resource, Api
import oms

app = Flask(__name__)
api = Api(app)

class HelloWorld(Resource):
	def get(self):
		return {'hello': 'world'}
class ExecuteOMS(Resource):
	def get(self):
		oms.CallOMS("Input/shipping_Falabella.xlsx","Plantilla/propuesta.xlsx")
		return{'message':'Generación exitosa'}
api.add_resource(HelloWorld, '/')
api.add_resource(ExecuteOMS,'/api/v1/OMS/ShippingMatrix')

if __name__ == '__main__':
	app.run(debug=True)
