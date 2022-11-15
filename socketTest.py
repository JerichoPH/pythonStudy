import socket

socket_client = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
a = socket_client.connect(('192.168.0.104', 8001))
print('-' * 10, a)
