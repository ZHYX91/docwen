"""
有效期检查模块（重构版）

本模块仅负责基于内置的、混淆的日期来检查程序的有效期状态。
不处理任何文件IO（如license.dat），也不执行任何UI操作（如弹窗）或退出操作。
"""

import base64
import datetime
import logging
from enum import Enum, auto

logger = logging.getLogger(__name__)

# --- 核心常量 ---
_OBFUSCATION_KEY = "n_V@c#2&pZ!q*S$r"
_OBFUSCATED_DATE = 'IzUXOS1KAhQ8DmRM' # 来自 scripts/maintenance/update_expiration.py

class ExpirationStatus(Enum):
    """定义了软件的有效期状态"""
    VALID = auto()                # 正常有效
    NEARING_EXPIRATION = auto()   # 临近过期
    EXPIRED = auto()              # 已过期
    TAMPERED = auto()             # 系统时间被篡改

class ExpirationInfo:
    """用于封装有效期检查结果的数据类"""
    def __init__(self, status: ExpirationStatus, days_left: int = 999):
        self.status = status
        self.days_left = days_left

    def __repr__(self):
        return f"<ExpirationInfo(status={self.status.name}, days_left={self.days_left})>"

def _deobfuscate(obfuscated_data: str, key: str) -> str:
    """
    执行XOR和Base64解码来反混淆数据。
    """
    try:
        decoded_b64 = base64.b64decode(obfuscated_data.encode('utf-8')).decode('utf-8')
        key_repeated = key * (len(decoded_b64) // len(key) + 1)
        xored = ''.join(chr(ord(c) ^ ord(k)) for c, k in zip(decoded_b64, key_repeated))
        return base64.b64decode(xored.encode('utf-8')).decode('utf-8')
    except Exception:
        return "2000-01-01"

def _get_expiration_date() -> datetime.datetime:
    """解密并返回硬编码的过期日期"""
    date_str = _deobfuscate(_OBFUSCATED_DATE, _OBFUSCATION_KEY)
    return datetime.datetime.strptime(date_str, '%Y-%m-%d')

# 在模块加载时计算一次，作为全局常量
EXPIRATION_DATE = _get_expiration_date()

def get_expiration_status() -> ExpirationInfo:
    """
    核心函数：检查当前有效期状态，不产生任何副作用。

    返回:
        ExpirationInfo: 包含当前状态和剩余天数的结果对象。
    """
    try:
        current_time = datetime.datetime.now()
        
        # 核心逻辑：比较当前时间和内置过期时间
        time_left = EXPIRATION_DATE - current_time
        days_left = time_left.days
        
        if days_left < 0:
            logger.warning(f"软件已于 {EXPIRATION_DATE.strftime('%Y-%m-%d')} 过期。")
            return ExpirationInfo(ExpirationStatus.EXPIRED, days_left)
        
        if days_left < 90: # 小于90天（不到3个月）
            logger.info(f"软件有效期不到3个月，剩余 {days_left} 天。")
            return ExpirationInfo(ExpirationStatus.NEARING_EXPIRATION, days_left)
            
        logger.debug("软件在有效期内。")
        return ExpirationInfo(ExpirationStatus.VALID, days_left)

    except Exception as e:
        logger.error(f"计算有效期状态时发生未知错误: {e}", exc_info=True)
        # 发生错误时，为安全起见，视为已过期
        return ExpirationInfo(ExpirationStatus.EXPIRED, -999)

def check_expiration():
    """
    旧的兼容性函数，用于启动时的临时检查。
    在未来将被完全移除，由 get_expiration_status 和 UI 逻辑替代。
    
    注意：此函数仍会调用 sys.exit()，但仅在检测到时间篡改时。
    """
    status_info = get_expiration_status()
    # 当前版本暂时不执行任何操作，仅打印日志
    logger.debug(f"启动时有效期检查完成: {status_info}")

# --- CLI 测试入口 ---
if __name__ == "__main__":
    print("测试有效期检查模块（重构版）...")
    
    # 场景1：正常有效期
    EXPIRATION_DATE = datetime.datetime.now() + datetime.timedelta(days=100)
    info = get_expiration_status()
    print(f"\n设置为远期过期 ({EXPIRATION_DATE.strftime('%Y-%m-%d')}):")
    print(f"  -> 结果: {info}")
    assert info.status == ExpirationStatus.VALID

    # 场景2：即将过期
    EXPIRATION_DATE = datetime.datetime.now() + datetime.timedelta(days=15)
    info = get_expiration_status()
    print(f"\n设置为即将过期 ({EXPIRATION_DATE.strftime('%Y-%m-%d')}):")
    print(f"  -> 结果: {info}")
    assert info.status == ExpirationStatus.NEARING_EXPIRATION
    
    # 场景3：已过期
    EXPIRATION_DATE = datetime.datetime.now() - datetime.timedelta(days=1)
    info = get_expiration_status()
    print(f"\n设置为已过期 ({EXPIRATION_DATE.strftime('%Y-%m-%d')}):")
    print(f"  -> 结果: {info}")
    assert info.status == ExpirationStatus.EXPIRED
    
    print("\n所有测试场景通过！")
