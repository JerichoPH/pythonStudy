import argparse
import socket

parser = argparse.ArgumentParser()
parser.description = 'socket链接工具'
parser.add_argument('-I', '--ip', help='监听IP，默认：127.0.0.1', type=str, default='127.0.0.1')
parser.add_argument('-P', '--port', help='监听端口，默认：8001', type=int, default=8001)
args = parser.parse_args()
ip = args.ip
port = args.port

sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
sock.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
sock.bind((ip, port,))  # 设置监听

sock.listen(5)  # 开启监听

while True:
    conn, addr = sock.accept()
    print(f'链接成功：{addr}')

    conn.sendall('链接成功'.encode('utf-8'))

    while True:
        # 接收消息
        data = conn.recv(1024)
        if not data:
            break
        data_string = data.decode('utf-8')

        #  回复消息
        conn.sendall(f'已接收消息：{data_string}'.encode('utf-8'))

    print('断开链接')

    # 关闭链接
    conn.close()
