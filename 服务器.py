from fastapi import FastAPI, HTTPException, Query  # pyright: ignore[reportMissingImports]
from fastapi.middleware.cors import CORSMiddleware  # pyright: ignore[reportMissingImports]
from pydantic import BaseModel  # pyright: ignore[reportMissingImports]
from typing import Optional, List
from contextlib import asynccontextmanager
import asyncpg  # pyright: ignore[reportMissingImports]
import threading
import time
from datetime import datetime, timedelta, timezone
import asyncio
import signal
import os

# 北京时区（UTC+8）
BEIJING_TZ = timezone(timedelta(hours=8))

def get_beijing_now():
    """获取北京时间的当前时间"""
    return datetime.now(BEIJING_TZ)

def get_beijing_today():
    """获取北京时间的今天日期（date对象）"""
    return get_beijing_now().date()

# 数据库配置
DB_CONFIG = {
    "host": "localhost",
    "port": 5432,
    "database": "qn_stats",
    "user": "postgres",
    "password": "Hzbd88888888"
}

# 连接池
db_pool: Optional[asyncpg.Pool] = None
active_employees = {}
active_employees_lock = threading.Lock()  # 线程锁保护active_employees

# 消息推送系统（实时消息，不存储历史）
pending_messages = {}  # {employee_id: [{"message": "xxx", "timestamp": xxx}, ...]}
pending_messages_lock = threading.Lock()
message_events = {}  # {employee_id: asyncio.Event()}，用于长轮询唤醒
message_events_lock = threading.Lock()

# Pydantic模型
class ReportData(BaseModel):
    employee_name: str
    report_date: Optional[str] = None  # 数据日期标记（格式：YYYY-MM-DD）
    report_timestamp: Optional[float] = None  # 数据上报时间戳
    total_customers: int = 0
    total_shops: int = 0
    shops_list: List[str] = []
    today_consult: int = 0
    today_replied: int = 0  # 新增：今日回复数
    total_reply_time: float = 0.0  # 新增：总回复时长
    avg_reply: int = 0
    online: bool = True

# UpdateStatsData 模型已废弃，不再使用

class RenameEmployee(BaseModel):
    original_id: str
    new_name: str

class DeleteEmployee(BaseModel):
    employee_id: str
    delete_all: bool = True  # True=删除所有记录, False=只删除今日记录

class ColorConfig(BaseModel):
    employee_id: str
    bar_color: Optional[str] = None
    line_color: Optional[str] = None

class SaveColorsRequest(BaseModel):
    colors: List[ColorConfig]  # 员工颜色配置列表
    global_line_color: Optional[str] = None  # 全局线形图颜色

class EmployeeOrder(BaseModel):
    employee_id: str
    order: int  # 排序序号，数字小的排前面

class SaveOrderRequest(BaseModel):
    orders: List[EmployeeOrder]  # 员工排序列表

class EmployeeVisibility(BaseModel):
    employee_id: str
    hidden: bool  # True=隐藏, False=显示
    is_manual: bool = True  # True=手动设置, False=自动隐藏

class SaveVisibilityRequest(BaseModel):
    visibility: List[EmployeeVisibility]  # 员工可见性配置列表

class GlobalVisibilityMode(BaseModel):
    show_all: bool

class SendMessageRequest(BaseModel):
    employee_id: str  # 员工原始ID
    message: str  # 消息内容  # True=显示所有员工, False=应用隐藏规则

class EmployeeResponse(BaseModel):
    employee_name: str
    display_name: str
    total_customers: int
    total_shops: int
    shops_list: List[str]
    today_consult: int
    avg_reply: int
    online: bool

class HistoryRecord(BaseModel):
    date: str
    total_consult: int
    avg_reply: int

# 数据库初始化
async def init_db():
    global db_pool
    db_pool = await asyncpg.create_pool(**DB_CONFIG, min_size=5, max_size=20)
    
    async with db_pool.acquire() as conn:
        # 创建每日统计表
        await conn.execute('''
            CREATE TABLE IF NOT EXISTS daily_stats (
                employee_id TEXT NOT NULL,
                date DATE NOT NULL,
                total_consultations INTEGER DEFAULT 0,
                replied_count INTEGER DEFAULT 0,
                total_reply_time REAL DEFAULT 0.0,
                avg_reply REAL DEFAULT 0.0,
                PRIMARY KEY (employee_id, date)
            )
        ''')
        
        # 创建员工元数据表
        await conn.execute('''
            CREATE TABLE IF NOT EXISTS employee_meta (
                original_id TEXT PRIMARY KEY,
                display_name TEXT NOT NULL
            )
        ''')
        
        # 添加排序字段（如果不存在）
        try:
            await conn.execute('ALTER TABLE employee_meta ADD COLUMN IF NOT EXISTS sort_order INTEGER DEFAULT 9999')
        except Exception:
            pass  # 字段可能已存在
        
        # 创建员工可见性表（隐藏状态）
        await conn.execute('''
            CREATE TABLE IF NOT EXISTS employee_visibility (
                employee_id TEXT PRIMARY KEY,
                hidden BOOLEAN DEFAULT FALSE,
                is_manual BOOLEAN DEFAULT TRUE,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')
        
        # 创建全局显示模式表
        await conn.execute('''
            CREATE TABLE IF NOT EXISTS global_visibility_mode (
                id INTEGER PRIMARY KEY DEFAULT 1,
                show_all BOOLEAN DEFAULT FALSE,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')
        
        # 插入默认的全局显示模式（如果不存在）
        await conn.execute('''
            INSERT INTO global_visibility_mode (id, show_all)
            VALUES (1, FALSE)
            ON CONFLICT (id) DO NOTHING
        ''')
        
        # 创建索引
        await conn.execute('''
            CREATE INDEX IF NOT EXISTS idx_daily_stats_date 
            ON daily_stats(date)
        ''')
        await conn.execute('''
            CREATE INDEX IF NOT EXISTS idx_daily_stats_employee 
            ON daily_stats(employee_id)
        ''')

async def check_daily_reset_flag():
    """定期检查跨天重置标志文件，并执行数据库清理"""
    import tempfile
    flag_file = os.path.join(tempfile.gettempdir(), "qn_stats_daily_reset_flag.txt")
    
    while True:
        try:
            await asyncio.sleep(5)  # 每5秒检查一次
            
            if os.path.exists(flag_file):
                try:
                    with open(flag_file, "r", encoding="utf-8") as f:
                        date_str = f.read().strip()
                    
                    current_date = datetime.strptime(date_str, "%Y-%m-%d").date()
                    print(f"[DAILY_RESET] 检测到跨天标志文件，开始清理数据库: {current_date}")
                    
                    # 执行数据库清理
                    await clear_today_stats_async(current_date)
                    
                    # 删除标志文件
                    os.remove(flag_file)
                    print(f"[DAILY_RESET] 数据库清理完成，标志文件已删除")
                except Exception as e:
                    print(f"[DAILY_RESET] 处理标志文件失败: {e}")
                    # 删除损坏的标志文件
                    try:
                        os.remove(flag_file)
                    except:
                        pass
        except Exception as e:
            print(f"[DAILY_RESET] 检查标志文件出错: {e}")

@asynccontextmanager
async def lifespan(app: FastAPI):
    # 启动时执行
    await init_db()
    # 启动清理线程
    threading.Thread(target=cleanup_inactive_sync, daemon=True).start()
    # 启动每日重置线程
    threading.Thread(target=daily_reset_sync, daemon=True).start()
    # 启动跨天标志检查任务（在事件循环中）
    asyncio.create_task(check_daily_reset_flag())
    yield
    # 关闭时执行
    if db_pool:
        await db_pool.close()

app = FastAPI(title="客服监控系统API", lifespan=lifespan)

# CORS配置
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

def cleanup_inactive_sync():
    """清理1分钟无上报的员工"""
    while True:
        now = time.time()
        with active_employees_lock:
            inactive = [eid for eid, data in active_employees.items() 
                       if now - data.get("last_seen", 0) > 60]  # 1分钟
            for eid in inactive:
                if eid in active_employees:
                    del active_employees[eid]
                    print(f"[CLEANUP] 清理离线员工: {eid}")
        time.sleep(30)  # 每30秒检查一次

def daily_reset_sync():
    """每天00:00清空当天的咨询数据和平均回复（历史数据永久保存在数据库中）"""
    last_reset_date = None
    
    # 首次启动时，记录当前日期作为基准（不清空数据，因为可能是当天重启）
    last_reset_date = get_beijing_today()
    print(f"[DAILY_RESET] 服务启动，当前日期（北京时间）: {last_reset_date}")
    print(f"[DAILY_RESET] 不清空数据（可能是当天重启，保留员工端已累积的数据）")
    
    while True:
        try:
            # 每10秒检查一次是否跨天（更快响应）
            time.sleep(10)
            
            current_date = get_beijing_today()
            
            # 检查是否到了新的一天（00:00之后）
            if last_reset_date != current_date:
                print(f"[DAILY_RESET] 检测到日期变化: {last_reset_date} -> {current_date}")
                print(f"[DAILY_RESET] 开始归档 {last_reset_date} 的数据")
                
                # 1. 归档并清空内存中的统计数据
                with active_employees_lock:
                    for eid, data in list(active_employees.items()):
                        old_date_str = data.get("date")
                        
                        # 如果员工数据还是昨天的，需要归档
                        if old_date_str:
                            try:
                                old_date = datetime.strptime(old_date_str, "%Y-%m-%d").date()
                                if old_date < current_date:
                                    print(f"[归档] 员工 {eid} 昨天({old_date})数据: {data.get('today_consult', 0)}咨询")
                                    # 数据库中已经实时更新了，不需要额外操作
                            except:
                                pass
                        
                        # 清空当天的咨询数据和平均回复，保留其他数据（客户数、店铺数等）
                        data["date"] = str(current_date)
                        data["today_consult"] = 0
                        data["avg_reply"] = 0
                
                # 2. 清理工作（清除手动隐藏状态等）- 使用线程安全的标志文件方式
                try:
                    # 创建标志文件，让主事件循环处理数据库清理
                    import tempfile
                    flag_file = os.path.join(tempfile.gettempdir(), "qn_stats_daily_reset_flag.txt")
                    with open(flag_file, "w", encoding="utf-8") as f:
                        f.write(str(current_date))
                    print(f"[DAILY_RESET] 已创建清理标志文件: {flag_file}")
                except Exception as e:
                    print(f"[DAILY_RESET] 创建标志文件失败: {e}")
                
                print(f"[DAILY_RESET] {current_date} - 跨天处理完成，昨天（{last_reset_date}）的数据已永久保存，今天的数据正常接收中")
                
                last_reset_date = current_date
            
        except Exception as e:
            print(f"[DAILY_RESET] 错误: {e}")
            time.sleep(60)  # 出错后等待1分钟再重试

async def clear_today_stats_async(current_date):
    """跨天后的清理工作：清空数据库中今天的数据（避免残留昨天的数据）
    
    注意：只删除今天（current_date）的数据，历史数据（昨天及以前）不会受影响
    """
    try:
        if db_pool:
            async with db_pool.acquire() as conn:
                # 修复：清空数据库中今天的数据（如果存在）
                # 因为可能是昨天跨天瞬间写入的残留数据，会导致今天显示昨天的值
                # 注意：这里只删除 date = current_date（今天）的数据，不会影响历史数据
                
                # 先查询一下今天有多少条数据，用于日志
                today_count = await conn.fetchval(
                    'SELECT COUNT(*) FROM daily_stats WHERE date = $1',
                    current_date
                )
                
                if today_count and today_count > 0:
                    print(f"[DAILY_RESET] 检测到今天（{current_date}）有 {today_count} 条数据，开始清空...")
                    
                    # 安全删除：只删除今天的数据，使用参数化查询防止SQL注入，并且确保不会误删历史数据
                    deleted_today = await conn.execute(
                        'DELETE FROM daily_stats WHERE date = $1',
                        current_date
                    )
                    print(f"[DAILY_RESET] {current_date} - 已清空数据库中今天的数据（共 {today_count} 条），历史数据未受影响")
                else:
                    print(f"[DAILY_RESET] {current_date} - 今天没有数据需要清空")
                
                # 清除所有手动隐藏状态（每天重置）
                deleted_count = await conn.fetchval(
                    'DELETE FROM employee_visibility WHERE is_manual = TRUE RETURNING COUNT(*)'
                )
                if deleted_count:
                    print(f"[DAILY_RESET] 已清除 {deleted_count} 个手动隐藏状态")
                
                # 验证历史数据是否保留（查询昨天及以前的数据数量）
                yesterday = current_date - timedelta(days=1)
                history_count = await conn.fetchval(
                    'SELECT COUNT(*) FROM daily_stats WHERE date < $1',
                    current_date
                )
                if history_count is not None:
                    print(f"[DAILY_RESET] ✓ 历史数据验证：昨天及以前的数据已保留（共 {history_count} 条记录）")
                
                print(f"[DAILY_RESET] 跨天完成，历史数据已永久保存在数据库中，今天的数据将从0开始累积")
    except Exception as e:
        print(f"[DAILY_RESET] 清理失败: {e}")
        import traceback
        traceback.print_exc()

@app.post("/clear_today_manual")
async def clear_today_manual():
    """手动清空今天的数据（用于补清今天00:00未清空的数据）"""
    try:
        today = get_beijing_today()
        
        # 1. 清空内存中的统计数据（强制清空所有员工，不管是否在线）
        with active_employees_lock:
            cleared_count = 0
            total_employees = len(active_employees)
            for eid, data in active_employees.items():
                # 强制清空，不管当前值是多少
                old_consult = data.get("today_consult", 0)
                old_reply = data.get("avg_reply", 0)
                data["today_consult"] = 0
                data["avg_reply"] = 0
                if old_consult > 0 or old_reply > 0:
                    cleared_count += 1
            print(f"[MANUAL_CLEAR] 已清空内存中 {cleared_count}/{total_employees} 个员工的当天统计数据")
        
        # 2. 清空数据库中的当天数据
        if db_pool:
            async with db_pool.acquire() as conn:
                result = await conn.execute(
                    'DELETE FROM daily_stats WHERE date = $1',
                    today
                )
                print(f"[MANUAL_CLEAR] 已清空数据库中的当天统计数据: {result}")
        
        return {
            "success": True,
            "message": f"已清空今天（{today}）的数据，历史数据已永久保存。注意：如果员工端还在运行，它们会继续上报数据，建议重启员工端。",
            "cleared_memory": cleared_count,
            "total_employees": total_employees
        }
    except Exception as e:
        print(f"[MANUAL_CLEAR] 错误: {e}")
        raise HTTPException(status_code=500, detail=f"清空失败: {str(e)}")

@app.post("/report")
async def receive_report(data: ReportData):
    """接收员工上报数据"""
    name = data.employee_name
    now = time.time()
    server_today = get_beijing_today()
    
    # 检查员工端发送的数据日期
    if data.report_date:
        try:
            report_date = datetime.strptime(data.report_date, "%Y-%m-%d").date()
        except ValueError:
            report_date = server_today
    else:
        # 如果员工端没有发送日期标记（旧版本），使用服务器今天的日期
        report_date = server_today
    
    # 检查数据时间戳（防止接收过旧的数据）
    if data.report_timestamp:
        report_time = datetime.fromtimestamp(data.report_timestamp, BEIJING_TZ)
        time_diff = (get_beijing_now() - report_time).total_seconds()
        if time_diff > 600:  # 超过10分钟
            print(f"[REPORT] 拒绝过旧的数据：员工 {name}，时间差 {time_diff}秒")
            return {"status": "rejected", "message": "数据时间戳过旧"}
    
    # 检查是否跨天（员工端发送的是昨天的数据）
    # 注意：如果是23:59:58上报但00:00:02才到达，数据时间戳在合理范围内应该接受
    if report_date < server_today:
        # 检查时间戳，如果是跨天瞬间的数据（时间差<30秒），应该接受
        if data.report_timestamp:
            time_diff = (get_beijing_now() - datetime.fromtimestamp(data.report_timestamp, BEIJING_TZ)).total_seconds()
            if time_diff < 30:  # 30秒内的数据认为是跨天瞬间的，接受并归入昨天
                print(f"[REPORT] 接收跨天瞬间的数据：员工 {name}，日期 {report_date}，时间差 {time_diff:.1f}秒")
                # 继续处理，归入昨天的数据
            else:
                print(f"[REPORT] 拒绝过旧的跨天数据：员工 {name}，日期 {report_date}，时间差 {time_diff:.1f}秒")
                return {"status": "rejected", "message": f"数据日期过旧（{report_date}），请发送今天（{server_today}）的数据"}
        else:
            # 没有时间戳，直接拒绝
            print(f"[REPORT] 拒绝昨天的数据（无时间戳）：员工 {name}，日期 {report_date}")
            return {"status": "rejected", "message": f"数据日期过旧（{report_date}），请发送今天（{server_today}）的数据"}
    
    # 正常接收员工端上报的数据
    final_consult = data.today_consult
    
    # 计算平均回复时长（服务器端重新计算，确保准确）
    if data.today_replied > 0:
        final_avg_reply = int(data.total_reply_time / data.today_replied)
    else:
        final_avg_reply = 0
    
    # 使用锁保护内存数据更新
    with active_employees_lock:
        previous_data = active_employees.get(name, {})
        data_changed = (
            previous_data.get("total_customers") != data.total_customers or
            previous_data.get("total_shops") != data.total_shops or
            previous_data.get("today_consult") != final_consult or
            previous_data.get("avg_reply") != final_avg_reply or
            previous_data.get("shops_list") != data.shops_list
        )
        
        # 更新内存中的活跃员工
        active_employees[name] = {
            "employee_name": name,
            "date": data.report_date,  # 记录数据日期
            "total_customers": data.total_customers,
            "total_shops": data.total_shops,
            "shops_list": data.shops_list,
            "today_consult": final_consult,
            "avg_reply": final_avg_reply,
            "online": True,
            "last_seen": now,
            "data_changed": data_changed
        }
        print(f"[REPORT] 员工 {name} 上报数据，日期: {report_date}, 咨询量: {final_consult}, 平均回复: {final_avg_reply}秒")
    
    # 实时更新数据库
    # 修复：如果日期是今天，直接覆盖（不使用GREATEST），确保跨天后今天的数据从0开始
    # 历史数据（昨天及以前）使用GREATEST取最大值，避免网络延迟导致的数据丢失
    try:
        async with db_pool.acquire() as conn:
            if report_date == server_today:
                # 今天的数据：直接覆盖（不使用GREATEST），确保跨天后重置为0
                await conn.execute('''
                    INSERT INTO daily_stats 
                        (employee_id, date, total_consultations, replied_count, total_reply_time, avg_reply)
                    VALUES ($1, $2, $3, $4, $5, $6)
                    ON CONFLICT (employee_id, date) 
                    DO UPDATE SET 
                        total_consultations = EXCLUDED.total_consultations,
                        replied_count = EXCLUDED.replied_count,
                        total_reply_time = EXCLUDED.total_reply_time,
                        avg_reply = EXCLUDED.avg_reply
                ''', name, report_date, final_consult, data.today_replied, data.total_reply_time, final_avg_reply)
            else:
                # 历史数据：使用GREATEST取最大值，避免网络延迟导致的数据丢失
                await conn.execute('''
                    INSERT INTO daily_stats 
                        (employee_id, date, total_consultations, replied_count, total_reply_time, avg_reply)
                    VALUES ($1, $2, $3, $4, $5, $6)
                    ON CONFLICT (employee_id, date) 
                    DO UPDATE SET 
                        total_consultations = GREATEST(daily_stats.total_consultations, EXCLUDED.total_consultations),
                        replied_count = GREATEST(daily_stats.replied_count, EXCLUDED.replied_count),
                        total_reply_time = GREATEST(daily_stats.total_reply_time, EXCLUDED.total_reply_time),
                        avg_reply = EXCLUDED.avg_reply
                ''', name, report_date, final_consult, data.today_replied, data.total_reply_time, final_avg_reply)
    except Exception as e:
        print(f"DB write error: {e}")
    
    return {"status": "ok"}

# /update_stats 接口已删除，功能已合并到 /report 接口

@app.get("/get_stats")
async def get_stats(employee_name: str = Query(..., description="员工名称")):
    """获取员工当天的统计数据"""
    today = get_beijing_today()
    
    try:
        async with db_pool.acquire() as conn:
            row = await conn.fetchrow('''
                SELECT total_consultations, replied_count, total_reply_time, avg_reply, date
                FROM daily_stats
                WHERE employee_id = $1 AND date = $2
            ''', employee_name, today)
            
            if row:
                return {
                    "data_date": str(row['date']),  # 返回数据日期
                    "today_consult": row['total_consultations'],
                    "replied_count": row['replied_count'],
                    "total_reply_time": float(row['total_reply_time']),
                    "avg_reply": int(row['avg_reply'])
                }
            else:
                return {
                    "data_date": str(today),  # 返回今天的日期
                    "today_consult": 0,
                    "replied_count": 0,
                    "total_reply_time": 0.0,
                    "avg_reply": 0
                }
    except Exception as e:
        print(f"DB get_stats error: {e}")
        return {
            "data_date": str(today),
            "today_consult": 0,
            "replied_count": 0,
            "total_reply_time": 0.0,
            "avg_reply": 0
        }

@app.post("/rename_employee")
async def rename_employee(data: RenameEmployee):
    """重命名员工"""
    if not data.original_id or not data.new_name.strip():
        raise HTTPException(status_code=400, detail="无效参数")
    
    try:
        async with db_pool.acquire() as conn:
            await conn.execute('''
                INSERT INTO employee_meta (original_id, display_name)
                VALUES ($1, $2)
                ON CONFLICT (original_id) 
                DO UPDATE SET display_name = EXCLUDED.display_name
            ''', data.original_id, data.new_name.strip())
    except Exception as e:
        print(f"DB rename error: {e}")
        raise HTTPException(status_code=500, detail="数据库错误")
    
    return {"success": True}

@app.post("/delete_employee")
async def delete_employee(data: DeleteEmployee):
    """删除员工记录"""
    if not data.employee_id or not data.employee_id.strip():
        raise HTTPException(status_code=400, detail="无效参数")
    
    try:
        async with db_pool.acquire() as conn:
            if data.delete_all:
                # 删除所有记录
                await conn.execute(
                    'DELETE FROM daily_stats WHERE employee_id = $1',
                    data.employee_id.strip()
                )
                await conn.execute(
                    'DELETE FROM employee_meta WHERE original_id = $1',
                    data.employee_id.strip()
                )
            else:
                # 只删除今日记录
                today = get_beijing_today()
                await conn.execute(
                    'DELETE FROM daily_stats WHERE employee_id = $1 AND date = $2',
                    data.employee_id.strip(), today
                )
                # 注意：不删除employee_meta，保留元数据
    except Exception as e:
        print(f"DB delete error: {e}")
        raise HTTPException(status_code=500, detail=f"数据库错误: {str(e)}")
    
    return {"success": True, "message": "删除成功"}

@app.get("/color_configs")
async def get_color_configs(employee_ids: Optional[str] = Query(default=None, description="员工ID列表，逗号分隔")):
    """获取员工颜色配置"""
    try:
        async with db_pool.acquire() as conn:
            # 确保字段存在
            try:
                await conn.execute('ALTER TABLE employee_meta ADD COLUMN IF NOT EXISTS bar_color TEXT')
                await conn.execute('ALTER TABLE employee_meta ADD COLUMN IF NOT EXISTS line_color TEXT')
            except Exception:
                pass  # 字段已存在
            
            # 获取全局线形图颜色
            global_record = await conn.fetchrow(
                'SELECT line_color FROM employee_meta WHERE original_id = $1',
                '__global__'
            )
            global_line_color = global_record['line_color'] if global_record and global_record.get('line_color') else '#2196F3'
            
            result = {"global_line_color": global_line_color, "employee_colors": {}}
            
            # 如果指定了员工ID列表，只查询这些员工
            if employee_ids:
                employee_id_list = [eid.strip() for eid in employee_ids.split(',') if eid.strip()]
                if employee_id_list:
                    records = await conn.fetch(
                        'SELECT original_id, bar_color FROM employee_meta WHERE original_id = ANY($1::text[])',
                        employee_id_list
                    )
                    for row in records:
                        result["employee_colors"][row['original_id']] = {
                            "bar_color": row.get('bar_color') or '#4CAF50'
                        }
            else:
                # 查询所有员工的颜色配置
                records = await conn.fetch('SELECT original_id, bar_color FROM employee_meta WHERE original_id != $1', '__global__')
                for row in records:
                    result["employee_colors"][row['original_id']] = {
                        "bar_color": row.get('bar_color') or '#4CAF50'
                    }
            
            return result
    except Exception as e:
        print(f"DB get_color_configs error: {e}")
        return {"global_line_color": "#2196F3", "employee_colors": {}}

@app.post("/save_color_configs")
async def save_color_configs(data: SaveColorsRequest):
    """保存颜色配置"""
    try:
        async with db_pool.acquire() as conn:
            # 确保字段存在
            try:
                await conn.execute('ALTER TABLE employee_meta ADD COLUMN IF NOT EXISTS bar_color TEXT')
                await conn.execute('ALTER TABLE employee_meta ADD COLUMN IF NOT EXISTS line_color TEXT')
            except Exception:
                pass  # 字段已存在
            
            # 保存全局线形图颜色
            if data.global_line_color:
                exists_global = await conn.fetchrow(
                    'SELECT original_id FROM employee_meta WHERE original_id = $1',
                    '__global__'
                )
                if exists_global:
                    await conn.execute(
                        'UPDATE employee_meta SET line_color = $1 WHERE original_id = $2',
                        data.global_line_color, '__global__'
                    )
                else:
                    await conn.execute(
                        'INSERT INTO employee_meta (original_id, display_name, line_color) VALUES ($1, $2, $3)',
                        '__global__', '全局设置', data.global_line_color
                    )
            
            # 保存每个员工的颜色配置
            for color_config in data.colors:
                if not color_config.employee_id or color_config.employee_id == '__global__':
                    continue
                
                exists = await conn.fetchrow(
                    'SELECT original_id FROM employee_meta WHERE original_id = $1',
                    color_config.employee_id
                )
                
                if exists:
                    # 更新
                    if color_config.bar_color:
                        await conn.execute(
                            'UPDATE employee_meta SET bar_color = $1 WHERE original_id = $2',
                            color_config.bar_color, color_config.employee_id
                        )
                else:
                    # 插入
                    await conn.execute(
                        'INSERT INTO employee_meta (original_id, display_name, bar_color) VALUES ($1, $2, $3)',
                        color_config.employee_id, color_config.employee_id, color_config.bar_color or '#4CAF50'
                    )
            
            return {"success": True, "message": "颜色配置已保存"}
    except Exception as e:
        print(f"DB save_color_configs error: {e}")
        raise HTTPException(status_code=500, detail=f"数据库错误: {str(e)}")

@app.get("/employee_order")
async def get_employee_order():
    """获取员工排序配置"""
    try:
        async with db_pool.acquire() as conn:
            # 确保字段存在
            try:
                await conn.execute('ALTER TABLE employee_meta ADD COLUMN IF NOT EXISTS sort_order INTEGER DEFAULT 9999')
            except Exception:
                pass
            
            records = await conn.fetch('SELECT original_id, sort_order FROM employee_meta WHERE sort_order IS NOT NULL ORDER BY sort_order')
            result = [{"employee_id": row['original_id'], "order": row['sort_order']} for row in records]
            return result
    except Exception as e:
        print(f"DB get_employee_order error: {e}")
        return []

@app.post("/save_employee_order")
async def save_employee_order(data: SaveOrderRequest):
    """保存员工排序配置"""
    try:
        async with db_pool.acquire() as conn:
            # 确保字段存在
            try:
                await conn.execute('ALTER TABLE employee_meta ADD COLUMN IF NOT EXISTS sort_order INTEGER DEFAULT 9999')
            except Exception:
                pass
            
            # 保存每个员工的排序
            for order_info in data.orders:
                exists = await conn.fetchrow(
                    'SELECT original_id FROM employee_meta WHERE original_id = $1',
                    order_info.employee_id
                )
                
                if exists:
                    # 更新排序
                    await conn.execute(
                        'UPDATE employee_meta SET sort_order = $1 WHERE original_id = $2',
                        order_info.order, order_info.employee_id
                    )
                else:
                    # 插入（如果没有记录，使用employee_id作为display_name）
                    await conn.execute(
                        'INSERT INTO employee_meta (original_id, display_name, sort_order) VALUES ($1, $2, $3)',
                        order_info.employee_id, order_info.employee_id, order_info.order
                    )
            
            return {"success": True, "message": "排序配置已保存"}
    except Exception as e:
        print(f"DB save_employee_order error: {e}")
        raise HTTPException(status_code=500, detail=f"数据库错误: {str(e)}")

@app.get("/employees", response_model=List[EmployeeResponse])
async def get_employees():
    """获取所有员工信息"""
    try:
        async with db_pool.acquire() as conn:
            # 获取今日所有有记录的员工
            today = get_beijing_today()  # 使用北京时间的date对象
            db_records = await conn.fetch('''
                SELECT employee_id, total_consultations, avg_reply 
                FROM daily_stats 
                WHERE date = $1
            ''', today)
            db_records_dict = {
                row['employee_id']: {
                    "consult": row['total_consultations'], 
                    "avg": int(row['avg_reply'])
                } 
                for row in db_records
            }
            
            # 获取自定义名称
            name_records = await conn.fetch('SELECT original_id, display_name FROM employee_meta')
            name_map = {row['original_id']: row['display_name'] for row in name_records}
        
        result = []
        now = time.time()
        
        # 添加活跃员工（使用锁保护）
        with active_employees_lock:
            employees_copy = dict(active_employees)  # 复制一份避免长时间持有锁
        
        for name, data in employees_copy.items():
            # 判断是否在线：1分钟内有上报即认为在线（不要求数据变化）
            time_since_last_seen = now - data.get("last_seen", 0)
            is_online = time_since_last_seen <= 60
            
            if not is_online:
                print(f"[GET_EMPLOYEES] 员工 {name} 显示为离线，距离上次上报 {time_since_last_seen:.1f} 秒")
            
            result.append(EmployeeResponse(
                employee_name=name,
                display_name=name_map.get(name, name),
                total_customers=data["total_customers"],
                total_shops=data["total_shops"],
                shops_list=data["shops_list"],
                today_consult=data["today_consult"],
                avg_reply=data["avg_reply"],
                online=is_online
            ))
            db_records_dict.pop(name, None)
        
        # 添加今日有数据但已离线的员工
        for name, stats in db_records_dict.items():
            result.append(EmployeeResponse(
                employee_name=name,
                display_name=name_map.get(name, name),
                total_customers=0,
                total_shops=0,
                shops_list=[],
                today_consult=stats["consult"],
                avg_reply=stats["avg"],
                online=False
            ))
        
        return result
        
    except Exception as e:
        print(f"DB read error in /employees: {e}")
        return []

@app.get("/history", response_model=List[HistoryRecord])
async def get_history(
    employee_id: str = Query(default="", description="员工ID"),
    period: str = Query(default="day", description="时间周期：day/week/month/custom"),
    start: Optional[str] = Query(default=None, description="开始日期"),
    end: Optional[str] = Query(default=None, description="结束日期")
):
    """查询历史统计"""
    try:
        async with db_pool.acquire() as conn:
            conditions = []
            params = []
            param_count = 1
            
            if employee_id:
                conditions.append(f"employee_id = ${param_count}")
                params.append(employee_id)
                param_count += 1
            
            if period == 'day':
                target = get_beijing_today()  # 使用北京时间的date对象
                conditions.append(f"date = ${param_count}")
                params.append(target)
                param_count += 1
            elif period == 'yesterday':
                # 新增：昨日
                yesterday = get_beijing_today() - timedelta(days=1)
                conditions.append(f"date = ${param_count}")
                params.append(yesterday)
                param_count += 1
            elif period == 'week':
                beijing_now = get_beijing_now()
                dates = [(beijing_now - timedelta(days=i)).date() for i in range(7)]  # 使用北京时间的date对象
                placeholders = ','.join([f'${i}' for i in range(param_count, param_count + len(dates))])
                conditions.append(f"date IN ({placeholders})")
                params.extend(dates)
                param_count += len(dates)
            elif period == 'month':
                beijing_now = get_beijing_now()
                month_start = beijing_now.replace(day=1).date()  # 使用北京时间的date对象
                next_month = (beijing_now.replace(day=1) + timedelta(days=32)).replace(day=1).date()
                conditions.append(f"date >= ${param_count} AND date < ${param_count + 1}")
                params.append(month_start)
                params.append(next_month)
                param_count += 2
            elif period == 'custom' and start and end:
                conditions.append(f"date BETWEEN ${param_count} AND ${param_count + 1}")
                # 将字符串转换为date对象
                params.extend([datetime.strptime(start, "%Y-%m-%d").date(), datetime.strptime(end, "%Y-%m-%d").date()])
                param_count += 2
            
            if conditions:
                query = f"SELECT date, total_consultations, avg_reply FROM daily_stats WHERE {' AND '.join(conditions)} ORDER BY date DESC"
            else:
                query = "SELECT date, total_consultations, avg_reply FROM daily_stats ORDER BY date DESC LIMIT 30"
            
            rows = await conn.fetch(query, *params)
            
            return [
                HistoryRecord(
                    date=str(row['date']),
                    total_consult=row['total_consultations'],
                    avg_reply=int(row['avg_reply'])
                )
                for row in rows
            ]
        
    except Exception as e:
        print(f"DB read error in /history: {e}")
        return []

@app.get("/stats_by_employee")
async def get_stats_by_employee(
    period: str = Query(default="day", description="时间周期：day/week/month/custom"),
    start: Optional[str] = Query(default=None, description="开始日期"),
    end: Optional[str] = Query(default=None, description="结束日期")
):
    """按员工分组统计（用于图表展示）"""
    try:
        async with db_pool.acquire() as conn:
            conditions = []
            params = []
            param_count = 1
            
            if period == 'day':
                target = get_beijing_today()
                conditions.append(f"date = ${param_count}")
                params.append(target)
                param_count += 1
            elif period == 'yesterday':
                # 新增：昨日
                yesterday = get_beijing_today() - timedelta(days=1)
                conditions.append(f"date = ${param_count}")
                params.append(yesterday)
                param_count += 1
            elif period == 'week':
                beijing_now = get_beijing_now()
                dates = [(beijing_now - timedelta(days=i)).date() for i in range(7)]
                placeholders = ','.join([f'${i}' for i in range(param_count, param_count + len(dates))])
                conditions.append(f"date IN ({placeholders})")
                params.extend(dates)
                param_count += len(dates)
            elif period == 'month':
                beijing_now = get_beijing_now()
                month_start = beijing_now.replace(day=1).date()
                next_month = (beijing_now.replace(day=1) + timedelta(days=32)).replace(day=1).date()
                conditions.append(f"date >= ${param_count} AND date < ${param_count + 1}")
                params.append(month_start)
                params.append(next_month)
                param_count += 2
            elif period == 'custom' and start and end:
                conditions.append(f"date BETWEEN ${param_count} AND ${param_count + 1}")
                params.extend([datetime.strptime(start, "%Y-%m-%d").date(), datetime.strptime(end, "%Y-%m-%d").date()])
                param_count += 2
            
            # 按员工分组聚合数据
            if conditions:
                query = f'''
                    SELECT 
                        employee_id,
                        SUM(total_consultations) as total_consult,
                        AVG(avg_reply) as avg_reply_time
                    FROM daily_stats
                    WHERE {' AND '.join(conditions)}
                    GROUP BY employee_id
                    ORDER BY total_consult DESC
                '''
            else:
                query = '''
                    SELECT 
                        employee_id,
                        SUM(total_consultations) as total_consult,
                        AVG(avg_reply) as avg_reply_time
                    FROM daily_stats
                    WHERE date >= CURRENT_DATE - INTERVAL '30 days'
                    GROUP BY employee_id
                    ORDER BY total_consult DESC
                '''
            
            rows = await conn.fetch(query, *params)
            
            # 获取自定义名称
            name_records = await conn.fetch('SELECT original_id, display_name FROM employee_meta')
            name_map = {row['original_id']: row['display_name'] for row in name_records}
            
            result = []
            for row in rows:
                employee_id = row['employee_id']
                total_consult = int(row['total_consult'] or 0)
                avg_reply = float(row['avg_reply_time'] or 0)
                
                # 计算效率指标（咨询数量 ÷ 平均回复时长，如果时长为0或咨询为0则返回0）
                if avg_reply > 0 and total_consult > 0:
                    efficiency = total_consult / avg_reply
                else:
                    efficiency = 0.0
                
                result.append({
                    "employee_name": name_map.get(employee_id, employee_id),
                    "employee_id": employee_id,
                    "total_consult": total_consult,
                    "avg_reply": int(avg_reply),
                    "efficiency": round(efficiency, 2)
                })
            
            return result
            
    except Exception as e:
        print(f"DB stats_by_employee error: {e}")
        return []

@app.get("/monthly_daily_stats")
async def get_monthly_daily_stats(
    year: Optional[int] = Query(default=None, description="年份，默认当前年"),
    month: Optional[int] = Query(default=None, description="月份，默认当前月")
):
    """获取指定月份所有员工的每日数据（用于客户端预加载和快速切换）"""
    try:
        # 如果未指定年月，使用当前年月（北京时间）
        beijing_now = get_beijing_now()
        target_year = year if year else beijing_now.year
        target_month = month if month else beijing_now.month
        
        # 计算月份的开始和结束日期
        month_start = datetime(target_year, target_month, 1).date()
        if target_month == 12:
            next_month = datetime(target_year + 1, 1, 1).date()
        else:
            next_month = datetime(target_year, target_month + 1, 1).date()
        
        async with db_pool.acquire() as conn:
            # 查询该月份所有员工的每日数据
            rows = await conn.fetch('''
                SELECT 
                    employee_id,
                    date,
                    total_consultations,
                    avg_reply
                FROM daily_stats
                WHERE date >= $1 AND date < $2
                ORDER BY employee_id, date
            ''', month_start, next_month)
            
            # 获取自定义名称
            name_records = await conn.fetch('SELECT original_id, display_name FROM employee_meta')
            name_map = {row['original_id']: row['display_name'] for row in name_records}
            
            # 按员工分组组织数据
            result = {}
            for row in rows:
                employee_id = row['employee_id']
                if employee_id not in result:
                    result[employee_id] = {
                        "employee_id": employee_id,
                        "employee_name": name_map.get(employee_id, employee_id),
                        "daily_data": []
                    }
                
                result[employee_id]["daily_data"].append({
                    "date": str(row['date']),
                    "total_consult": int(row['total_consultations'] or 0),
                    "avg_reply": int(row['avg_reply'] or 0)
                })
            
            # 转换为列表返回
            return {
                "year": target_year,
                "month": target_month,
                "month_start": str(month_start),
                "month_end": str(next_month - timedelta(days=1)),
                "employees": list(result.values())
            }
            
    except Exception as e:
        print(f"DB monthly_daily_stats error: {e}")
        beijing_now = get_beijing_now()
        return {
            "year": year or beijing_now.year,
            "month": month or beijing_now.month,
            "employees": []
        }

@app.get("/employee_visibility")
async def get_employee_visibility():
    """获取所有员工的可见性配置"""
    try:
        if db_pool:
            async with db_pool.acquire() as conn:
                rows = await conn.fetch('SELECT employee_id, hidden, is_manual FROM employee_visibility')
                return [{"employee_id": row['employee_id'], "hidden": row['hidden'], "is_manual": row['is_manual']} for row in rows]
        return []
    except Exception as e:
        print(f"获取员工可见性失败: {e}")
        return []

@app.post("/employee_visibility")
async def save_employee_visibility(request: SaveVisibilityRequest):
    """保存员工可见性配置"""
    try:
        if db_pool:
            async with db_pool.acquire() as conn:
                for item in request.visibility:
                    await conn.execute('''
                        INSERT INTO employee_visibility (employee_id, hidden, is_manual, updated_at)
                        VALUES ($1, $2, $3, CURRENT_TIMESTAMP)
                        ON CONFLICT (employee_id) 
                        DO UPDATE SET hidden = $2, is_manual = $3, updated_at = CURRENT_TIMESTAMP
                    ''', item.employee_id, item.hidden, item.is_manual)
                print(f"[可见性] 已保存 {len(request.visibility)} 个员工的可见性配置")
                return {"success": True, "message": f"已保存 {len(request.visibility)} 个配置"}
        return {"success": False, "message": "数据库连接失败"}
    except Exception as e:
        print(f"保存员工可见性失败: {e}")
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/global_visibility_mode")
async def get_global_visibility_mode():
    """获取全局显示模式"""
    try:
        if db_pool:
            async with db_pool.acquire() as conn:
                row = await conn.fetchrow('SELECT show_all FROM global_visibility_mode WHERE id = 1')
                if row:
                    return {"show_all": row['show_all']}
        return {"show_all": False}
    except Exception as e:
        print(f"获取全局显示模式失败: {e}")
        return {"show_all": False}

@app.post("/global_visibility_mode")
async def save_global_visibility_mode(request: GlobalVisibilityMode):
    """保存全局显示模式"""
    try:
        if db_pool:
            async with db_pool.acquire() as conn:
                await conn.execute('''
                    UPDATE global_visibility_mode 
                    SET show_all = $1, updated_at = CURRENT_TIMESTAMP 
                    WHERE id = 1
                ''', request.show_all)
                mode_text = "显示所有员工" if request.show_all else "应用隐藏规则"
                print(f"[全局显示模式] 已更新为: {mode_text}")
                return {"success": True, "message": f"已更新为: {mode_text}"}
        return {"success": False, "message": "数据库连接失败"}
    except Exception as e:
        print(f"保存全局显示模式失败: {e}")
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/send_message")
async def send_message(request: SendMessageRequest):
    """管理端发送消息给员工端"""
    try:
        employee_id = request.employee_id
        message = request.message.strip()
        
        if not message:
            raise HTTPException(status_code=400, detail="消息内容不能为空")
        
        # 存储消息到待处理队列
        with pending_messages_lock:
            if employee_id not in pending_messages:
                pending_messages[employee_id] = []
            pending_messages[employee_id].append({
                "message": message,
                "timestamp": time.time()
            })
        
        # 唤醒长轮询（如果存在）
        with message_events_lock:
            if employee_id in message_events:
                message_events[employee_id].set()
        
        print(f"[消息推送] 发送给 {employee_id}: {message[:50]}...")
        return {"success": True, "message": "消息已发送"}
    except HTTPException:
        raise
    except Exception as e:
        print(f"发送消息失败: {e}")
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/poll_messages/{employee_id}")
async def poll_messages(employee_id: str):
    """员工端长轮询获取消息（30秒超时）"""
    try:
        # 检查是否有待处理消息
        with pending_messages_lock:
            if employee_id in pending_messages and pending_messages[employee_id]:
                messages = pending_messages[employee_id].copy()
                pending_messages[employee_id] = []
                return {"messages": messages}
        
        # 没有消息，创建事件等待
        with message_events_lock:
            if employee_id not in message_events:
                message_events[employee_id] = asyncio.Event()
            event = message_events[employee_id]
        
        # 等待30秒或直到有新消息
        try:
            await asyncio.wait_for(event.wait(), timeout=30.0)
            event.clear()
            
            # 再次检查消息
            with pending_messages_lock:
                if employee_id in pending_messages and pending_messages[employee_id]:
                    messages = pending_messages[employee_id].copy()
                    pending_messages[employee_id] = []
                    return {"messages": messages}
        except asyncio.TimeoutError:
            pass
        
        return {"messages": []}
    except Exception as e:
        print(f"长轮询错误: {e}")
        return {"messages": []}

@app.get("/")
async def root():
    return {"message": "客服监控系统API", "version": "2.0", "database": "PostgreSQL"}

if __name__ == "__main__":
    import uvicorn  # pyright: ignore[reportMissingImports]
    import sys
    
    # 定义清理函数
    def cleanup_handler(signum=None, frame=None):
        """信号处理器：确保资源被正确清理"""
        print("\n[SHUTDOWN] 收到退出信号，正在清理资源...")
        if db_pool:
            try:
                # 在新的事件循环中关闭连接池
                loop = asyncio.new_event_loop()
                asyncio.set_event_loop(loop)
                loop.run_until_complete(db_pool.close())
                loop.close()
                print("[SHUTDOWN] 数据库连接池已关闭")
            except Exception as e:
                print(f"[SHUTDOWN] 关闭数据库连接池失败: {e}")
        print("[SHUTDOWN] 清理完成，程序退出")
        sys.exit(0)
    
    # 注册信号处理器（Windows支持的信号有限）
    signal.signal(signal.SIGINT, cleanup_handler)  # Ctrl+C
    signal.signal(signal.SIGTERM, cleanup_handler)  # kill命令
    if hasattr(signal, 'SIGBREAK'):  # Windows特有
        signal.signal(signal.SIGBREAK, cleanup_handler)
    
    # Windows下双击运行时显示信息
    if sys.platform == "win32":
        os.system("title 客服监控系统 - FastAPI服务器")
        print("=" * 50)
        print("  客服监控系统 v2.0 - FastAPI服务器")
        print("=" * 50)
        print(f"  端口: 9999")
        print(f"  数据库: PostgreSQL")
        print(f"  地址: http://0.0.0.0:9999")
        print(f"  API文档: http://localhost:9999/docs")
        print("=" * 50)
        print("  服务器正在启动...")
        print("  按 Ctrl+C 停止服务器")
        print("=" * 50)
        print()
    
    try:
        uvicorn.run(app, host="0.0.0.0", port=9999)
    except KeyboardInterrupt:
        cleanup_handler()
    except Exception as e:
        print(f"[ERROR] 服务器异常: {e}")
        cleanup_handler()
