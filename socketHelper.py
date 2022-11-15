import argparse
import socket


class SocketClient:
    _socket_client = None
    _ip_addr: str = None
    _port_number: int = None

    def __init__(self, ip_addr: str, port_number: int):
        self._socket_client = socket.socket(socket.AF_INET, socket.SOCK_STREAM)

    @property
    def get_ip_addr(self) -> str:
        """
        获取IP地址
        :return: IP地址
        :rtype: str
        """
        return self._ip_addr

    def set_ip_addr(self, ip_addr: str) -> __init__:
        """
        设置IP地址
        :param ip_addr: IP地址
        :type ip_addr: str
        :return: 本类对象
        :rtype: SocketClient
        """
        self._ip_addr = ip_addr
        return self

    @property
    def get_port_number(self) -> int:
        """
        获取端口号
        :return: 端口号
        :rtype: int
        """
        return self._port_number

    def set_port_number(self, port_number: int) -> __init__:
        """
        设置端口号
        :param port_number: 端口号
        :type port_number: int
        :return: 本类对象
        :rtype: SocketClient
        """
        self._port_number = port_number
        return self

    def connect(self):
        self._socket_client.connect((self.get_ip_addr, self.get_port_number))


class SocketServer:
    _socket_server = None
    _ip_addr: str = None
    _port_number: int = None

    def __init__(self, ip_addr: str, port_number: int):
        """
        初始化
        :param ip_addr: ip地址
        :type ip_addr: str
        :param port_number: 端口号
        :type port_number: int
        """
        self._ip_addr = ip_addr
        self._port_number = port_number
        self._socket_server = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        self._socket_server.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
        self._socket_server.bind((self._ip_addr, self._port_number,))  # 设置监听

    def listen(self, num: int = 5, when_recv=None) -> None:
        """
        开启监听
        :param num: 连接数
        :type num: int
        :param when_recv: 接收信息时执行
        :type when_recv:
        """
        self._socket_server.listen(num)  # 开启监听
        print(f'开启监听{self._ip_addr}:{self._port_number}')
        while True:
            conn, addr = self._socket_server.accept()
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

            # print('断开链接')
            # conn.close()

    def get_socket(self) -> socket.socket:
        """
        获取socket对象
        :return: socket对象
        :rtype: socket.socket
        """
        return self._socket_server

    def close(self) -> None:
        """
        关闭链接
        :return: 无
        :rtype: None
        """
        self._socket_server.close()


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.description = 'socket链接工具'
    parser.add_argument('-I', '--ip', help='监听IP，默认：127.0.0.1', type=str, default='127.0.0.1')
    parser.add_argument('-P', '--port', help='监听端口，默认：8001', type=int, default=8001)
    args = parser.parse_args()
    ip = args.ip
    port = args.port


    def callback(conn, addr, data):
        data_string = data.decode('utf-8')
        conn.sendall(f'已接收消息：{data_string}'.encode('utf-8'))


    socket_server = SocketServer(ip_addr=ip, port_number=port)
    socket_server.listen(5, callback)
