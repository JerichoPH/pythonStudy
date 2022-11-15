import argparse
import socket

parser = argparse.ArgumentParser()
parser.description = 'socket链接工具'
parser.add_argument('-I', '--ip', help='监听IP，默认：127.0.0.1', type=str, default='127.0.0.1')
parser.add_argument('-P', '--port', help='监听端口，默认：8001', type=int, default=8001)
args = parser.parse_args()
ip = args.ip
port = args.port


class SocketServer:
    _socket = None

    def __init__(self, ip_addr: str, port_number: int):
        """
        初始化
        :param ip_addr: ip地址
        :type ip_addr: str
        :param port_number: 端口号
        :type port_number: int
        """
        self._socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        self._socket.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
        self._socket.bind((ip_addr, port_number,))  # 设置监听

    def listen(self, num: int = 5, when_recv=None) -> None:
        """
        开启监听
        :param num: 线程数
        :type num: int
        :param when_recv: 接收信息时执行
        :type when_recv:
        """
        self._socket.listen(num)  # 开启监听
        while True:
            conn, addr = self._socket.accept()
            print(f'链接成功：{addr}')

            conn.sendall('链接成功'.encode('utf-8'))

            while True:
                # 接收消息
                data = conn.recv(1024)
                if not data:
                    break
                when_recv(conn, addr, data)
                # data_string = data.decode('utf-8')

                #  回复消息
                # conn.sendall(f'已接收消息：{data_string}'.encode('utf-8'))

            print('断开链接')
            conn.close()

    def get_socket(self) -> socket.socket:
        """
        获取socket对象
        :return: socket对象
        :rtype: socket.socket
        """
        return self._socket

    def close(self) -> None:
        """
        关闭链接
        :return: 无
        :rtype: None
        """
        self._socket.close()


if __name__ == '__main__':
    def callback(conn, addr, data):
        data_string = data.decode('utf-8')
        conn.sendall(f'已接收消息：{data_string}'.encode('utf-8'))


    SocketServer(ip_addr=ip, port_number=port).listen(5, callback)
