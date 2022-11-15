import socket

from socketHelper import SocketClient

socket_client = SocketClient('127.0.0.1', 8001)
socket_client.connect()
