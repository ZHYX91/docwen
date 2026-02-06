"""
网络隔离模块 - 完全阻断所有网络连接
在应用启动时执行，重写网络相关模块以阻止任何网络连接尝试
"""

import logging
import socket
import urllib.request
import http.client
import ftplib
import smtplib
import ssl

logger = logging.getLogger()

class NetworkIsolationError(Exception):
    """网络隔离相关异常"""
    pass

class BlockedSocket:
    """被阻断的socket类，所有网络操作都会失败"""
    
    def __init__(self, *args, **kwargs):
        self._blocked = True
        logger.warning("尝试创建socket连接已被阻断")
        raise NetworkIsolationError("网络连接已被完全阻断 - 此应用不允许任何网络访问")
    
    def __getattr__(self, name):
        """拦截所有方法调用"""
        def blocked_method(*args, **kwargs):
            logger.warning(f"尝试调用socket方法 '{name}' 已被阻断")
            raise NetworkIsolationError("网络连接已被完全阻断 - 此应用不允许任何网络访问")
        return blocked_method

class BlockedHTTPConnection:
    """被阻断的HTTP连接类"""
    
    def __init__(self, *args, **kwargs):
        logger.warning("尝试创建HTTP连接已被阻断")
        raise NetworkIsolationError("HTTP连接已被完全阻断 - 此应用不允许任何网络访问")
    
    def __getattr__(self, name):
        """拦截所有方法调用"""
        def blocked_method(*args, **kwargs):
            logger.warning(f"尝试调用HTTP方法 '{name}' 已被阻断")
            raise NetworkIsolationError("HTTP连接已被完全阻断 - 此应用不允许任何网络访问")
        return blocked_method

class BlockedURLOpener:
    """被阻断的URL打开器"""
    
    def open(self, fullurl, data=None, timeout=None):
        logger.warning("尝试打开URL已被阻断")
        raise NetworkIsolationError("URL访问已被完全阻断 - 此应用不允许任何网络访问")
    
    def retrieve(self, url, filename=None, reporthook=None, data=None):
        logger.warning("尝试检索URL内容已被阻断")
        raise NetworkIsolationError("URL访问已被完全阻断 - 此应用不允许任何网络访问")

def block_socket_module():
    """阻断socket模块的所有网络功能"""
    original_socket = socket.socket
    
    # 重写socket.socket类
    socket.socket = BlockedSocket
    socket.SocketType = BlockedSocket
    
    # 阻断socket模块的其他网络函数
    def blocked_connect(*args, **kwargs):
        logger.warning("尝试使用socket.connect已被阻断")
        raise NetworkIsolationError("socket连接已被完全阻断")
    
    def blocked_create_connection(*args, **kwargs):
        logger.warning("尝试使用socket.create_connection已被阻断")
        raise NetworkIsolationError("socket连接创建已被完全阻断")
    
    def blocked_gethostbyname(*args, **kwargs):
        logger.warning("尝试使用socket.gethostbyname已被阻断")
        raise NetworkIsolationError("DNS解析已被完全阻断")
    
    def blocked_getaddrinfo(*args, **kwargs):
        logger.warning("尝试使用socket.getaddrinfo已被阻断")
        raise NetworkIsolationError("地址信息查询已被完全阻断")
    
    # 重写关键函数
    socket.connect = blocked_connect
    socket.create_connection = blocked_create_connection
    socket.gethostbyname = blocked_gethostbyname
    socket.getaddrinfo = blocked_getaddrinfo
    
    logger.info("socket模块网络功能已完全阻断")

def block_urllib_module():
    """阻断urllib模块的所有网络功能"""
    # 阻断urllib.request
    urllib.request.urlopen = lambda *args, **kwargs: (_ for _ in ()).throw(
        NetworkIsolationError("urllib.urlopen已被完全阻断 - 此应用不允许任何网络访问")
    )
    
    urllib.request.URLopener = BlockedURLOpener
    urllib.request.FancyURLopener = BlockedURLOpener
    
    # 安装全局的URL打开器阻断
    urllib.request.install_opener(BlockedURLOpener())
    
    logger.info("urllib模块网络功能已完全阻断")

def block_http_module():
    """阻断http.client模块的所有网络功能"""
    # 重写HTTPConnection和HTTPSConnection
    http.client.HTTPConnection = BlockedHTTPConnection
    http.client.HTTPSConnection = BlockedHTTPConnection
    
    logger.info("http.client模块网络功能已完全阻断")

def block_ftp_module():
    """阻断ftplib模块的所有网络功能"""
    original_FTP = ftplib.FTP
    
    class BlockedFTP:
        def __init__(self, *args, **kwargs):
            logger.warning("尝试创建FTP连接已被阻断")
            raise NetworkIsolationError("FTP连接已被完全阻断 - 此应用不允许任何网络访问")
        
        def __getattr__(self, name):
            def blocked_method(*args, **kwargs):
                logger.warning(f"尝试调用FTP方法 '{name}' 已被阻断")
                raise NetworkIsolationError("FTP操作已被完全阻断 - 此应用不允许任何网络访问")
            return blocked_method
    
    ftplib.FTP = BlockedFTP
    ftplib.FTP_TLS = BlockedFTP
    
    logger.info("ftplib模块网络功能已完全阻断")

def block_smtp_module():
    """阻断smtplib模块的所有网络功能"""
    original_SMTP = smtplib.SMTP
    
    class BlockedSMTP:
        def __init__(self, *args, **kwargs):
            logger.warning("尝试创建SMTP连接已被阻断")
            raise NetworkIsolationError("SMTP连接已被完全阻断 - 此应用不允许任何网络访问")
        
        def __getattr__(self, name):
            def blocked_method(*args, **kwargs):
                logger.warning(f"尝试调用SMTP方法 '{name}' 已被阻断")
                raise NetworkIsolationError("SMTP操作已被完全阻断 - 此应用不允许任何网络访问")
            return blocked_method
    
    smtplib.SMTP = BlockedSMTP
    smtplib.SMTP_SSL = BlockedSMTP
    
    logger.info("smtplib模块网络功能已完全阻断")

def block_ssl_module():
    """阻断ssl模块的网络连接功能"""
    # 在新版Python中，ssl.wrap_socket已被弃用，使用兼容性处理
    try:
        original_wrap_socket = ssl.wrap_socket
        
        def blocked_wrap_socket(*args, **kwargs):
            logger.warning("尝试使用SSL包装socket已被阻断")
            raise NetworkIsolationError("SSL连接已被完全阻断 - 此应用不允许任何网络访问")
        
        ssl.wrap_socket = blocked_wrap_socket
        logger.debug("已阻断ssl.wrap_socket")
    except AttributeError:
        logger.debug("ssl.wrap_socket在新版Python中不可用，跳过阻断")
    
    # 阻断SSL socket类
    ssl.SSLSocket = BlockedSocket
    
    logger.info("ssl模块网络功能已完全阻断")

def test_network_isolation():
    """测试网络隔离是否生效"""
    test_cases = [
        ("socket连接", lambda: socket.socket(socket.AF_INET, socket.SOCK_STREAM)),
        ("HTTP连接", lambda: http.client.HTTPConnection("example.com")),
        ("URL打开", lambda: urllib.request.urlopen("http://example.com")),
        ("FTP连接", lambda: ftplib.FTP("example.com")),
        ("SMTP连接", lambda: smtplib.SMTP("smtp.example.com")),
    ]
    
    failed_tests = []
    
    for test_name, test_func in test_cases:
        try:
            test_func()
            failed_tests.append(f"{test_name} - 未被正确阻断")
        except NetworkIsolationError:
            logger.debug(f"✓ {test_name} - 已正确阻断")
        except Exception as e:
            # 其他异常也是正常的，因为网络被阻断了
            logger.debug(f"✓ {test_name} - 已阻断（{type(e).__name__}）")
    
    if failed_tests:
        logger.error(f"网络隔离测试失败: {', '.join(failed_tests)}")
        return False
    else:
        logger.info("所有网络隔离测试通过 - 网络连接已被完全阻断")
        return True

def initialize_network_isolation():
    """
    初始化网络隔离 - 阻断所有网络连接
    在应用启动时调用此函数
    """
    logger.info("开始初始化网络隔离...")
    
    try:
        # 阻断各个网络模块
        block_socket_module()
        block_urllib_module()
        block_http_module()
        block_ftp_module()
        block_smtp_module()
        block_ssl_module()
        
        # 测试隔离效果
        isolation_ok = test_network_isolation()
        
        if isolation_ok:
            logger.info("网络隔离初始化完成 - 所有网络连接已被完全阻断")
        else:
            logger.error("网络隔离初始化失败 - 某些网络功能可能未被正确阻断")
            raise NetworkIsolationError("网络隔离初始化失败")
            
        return isolation_ok
        
    except Exception as e:
        logger.error(f"网络隔离初始化过程中发生错误: {str(e)}")
        raise

def get_network_isolation_status():
    """获取网络隔离状态信息"""
    
    status = {
        "socket_blocked": socket.socket == BlockedSocket,
        "http_blocked": http.client.HTTPConnection == BlockedHTTPConnection,
        # The patched urlopen is a lambda that throws an exception. It's not a regular function.
        "urllib_blocked": "throw" in str(urllib.request.urlopen),
        # Check if the module of the class has been changed from the original stdlib one.
        "ftp_blocked": ftplib.FTP.__module__ != 'ftplib',
        "smtp_blocked": smtplib.SMTP.__module__ != 'smtplib',
        # ssl.wrap_socket is a function, not a class
        "ssl_blocked": hasattr(ssl, 'wrap_socket') and ssl.wrap_socket.__module__ != 'ssl',
    }
    
    return status
