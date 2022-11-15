import socket

socket_client = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
a = socket_client.connect(('127.0.0.1', 8001))
print('-' * 10, a)
