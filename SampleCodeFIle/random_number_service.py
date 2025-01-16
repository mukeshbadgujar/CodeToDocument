import pika
import json
import random
from flask import Flask, request, jsonify

app = Flask(__name__)


@app.route('/generate', methods=['POST'])
def generate_random_number():
	data = request.json
	if 'username' not in data:
		return jsonify({'message': 'Username required'}), 400
	
	number = random.randint(1, 100)
	
	# Send message to RabbitMQ
	send_message_to_queue(data['username'], number)
	
	return jsonify({'message': f'Random number generated: {number}'}), 200


def send_message_to_queue(username, number):
	connection = pika.BlockingConnection(pika.ConnectionParameters('rabbitmq', 5672))
	channel = connection.channel()
	
	# Declare a queue
	channel.queue_declare(queue='random_number_queue')
	
	# Create the message
	message = json.dumps({'username': username, 'random_number': number})
	
	# Send the message
	channel.basic_publish(exchange='', routing_key='random_number_queue', body=message)
	
	connection.close()


if __name__ == '__main__':
	app.run(port=5002)
