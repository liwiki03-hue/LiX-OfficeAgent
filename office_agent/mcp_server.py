import argparse
import asyncio
import json
import logging
import os
import sys
import threading

from .mcp_host import McpHost, TOOLS


def create_server(host):
    from mcp.server.lowlevel import Server
    from mcp.types import Tool, ToolAnnotations, TextContent, CallToolResult

    server = Server('office-agent-local')
    busy = asyncio.Lock()

    @server.list_tools()
    async def list_tools():
        return [Tool(name=name, description=description, inputSchema={
            'type': 'object', 'properties': fields, 'required': list(fields), 'additionalProperties': False},
            annotations=ToolAnnotations(readOnlyHint=not name.startswith('plan_'), destructiveHint=False,
                                        idempotentHint=not name.startswith('plan_'), openWorldHint=False))
                for name, (description, fields) in TOOLS.items()]

    @server.call_tool(validate_input=False)
    async def call_tool(name, arguments):
        if busy.locked():
            result = {'ok': False, 'reason': 'busy'}
        else:
            async with busy:
                result = await asyncio.to_thread(host.call, name, arguments)
        return CallToolResult(content=[TextContent(type='text', text=json.dumps(result, ensure_ascii=True, allow_nan=False))],
                              isError=result.get('ok') is not True)

    return server


class BoundedInput:
    def __aiter__(self):
        return self

    async def __anext__(self):
        import anyio
        line = await anyio.to_thread.run_sync(sys.stdin.buffer.readline, 8193, abandon_on_cancel=True)
        if not line or len(line) > 8192:
            raise StopAsyncIteration
        try:
            return line.decode('utf8', errors='strict')
        except UnicodeError:
            raise StopAsyncIteration from None


async def serve(host):
    from mcp.server.stdio import stdio_server

    server = create_server(host)
    async with stdio_server(stdin=BoundedInput()) as (reader, writer):
        await server.run(reader, writer, server.create_initialization_options())


def control_window(host, done):
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk

    root = tk.Tk()
    root.title('LiX · MCP 本机授权')
    if os.name == 'nt':
        from .branding import APP_ICON
        root.iconbitmap(str(APP_ICON))
    root.geometry('620x700')
    root.minsize(480, 560)
    frame = ttk.Frame(root, padding=16)
    frame.pack(fill='both', expand=True)
    ttk.Label(frame, text='MCP 不是文件授权。先登录，再主动选择一个普通 DOCX/XLSX。', wraplength=550).pack(anchor='w')
    ttk.Label(frame, text='客户端只能查看计数和提出计划，不返回正文或路径。写入和撤销只在此窗口确认。文件许可15分钟，计划5分钟。', wraplength=550).pack(anchor='w', pady=8)
    ttk.Label(frame, text='用户名（请先在主程序注册本机账户）').pack(anchor='w')
    username = ttk.Entry(frame)
    username.pack(fill='x')
    ttk.Label(frame, text='密码').pack(anchor='w')
    password = ttk.Entry(frame, show='●')
    password.pack(fill='x')
    status = tk.StringVar(value='尚未登录；所有MCP业务调用都会拒绝。')
    ttk.Label(frame, textvariable=status, wraplength=550).pack(anchor='w', pady=8)
    plans = tk.Listbox(frame, height=7, exportselection=False)
    preview = tk.Text(frame, height=10, wrap='word', state='disabled')

    def say(value):
        status.set(value)

    def login():
        host.revoke()
        reply = host.auth.login(username.get(), password.get())
        password.delete(0, 'end')
        say('已登录，请主动选择文件；尚无文件许可。' if reply['authenticated'] else reply['message'])

    def choose():
        if host.disabled:
            say('本窗口有结果不确定的事务，请先核对备份和审计；操作保持关闭，重启后再授权。')
            return
        if host.auth.status()['authenticated'] is not True:
            say('请先登录。')
            return
        selected = filedialog.askopenfilename(parent=root, filetypes=[('Office 文档','*.docx *.xlsx')])
        if not selected:
            return
        if not messagebox.askyesno('文件许可', '允许本次MCP连接检查该文件并提出修改计划？\n只返回元数据；不会自动批准写入。重新选择会退休旧计划和撤销入口。', parent=root):
            return
        try:
            host.authorize(selected)
            say('文件许可已生效15分钟：' + host.path.name)
        except Exception:
            say('文件未获准；只支持通过现有路径和结构检查的普通DOCX/XLSX。')

    buttons = ttk.Frame(frame)
    buttons.pack(fill='x')
    ttk.Button(buttons, text='登录', command=login).pack(side='left')
    ttk.Button(buttons, text='选择并授权文件', command=choose).pack(side='left', padx=8)

    def revoke():
        host.logout()
        password.delete(0, 'end')
        say('已退出并清除文件许可、待确认计划和本窗口撤销入口。')

    ttk.Button(buttons, text='退出并撤销授权', command=revoke).pack(side='left')
    plans.pack(fill='x', pady=10)
    preview.pack(fill='both', expand=True)
    visible = []

    def selected_plan():
        selected = plans.curselection()
        if not selected or selected[0] >= len(visible):
            raise ValueError()
        return visible[selected[0]]

    def show(event=None):
        try:
            identifier = selected_plan()
            if identifier in host.history:
                text = '本机已确认写入副本；源文件保留。可在下方独立确认撤销副本。'
            else:
                value = host.preview(identifier)
                text = '本机预览（不返回MCP）\n' + value['target'] + '\n\n修改前：\n' + value['before'] + '\n\n修改后：\n' + value['after']
            preview.configure(state='normal')
            preview.delete('1.0', 'end')
            preview.insert('1.0', text)
            preview.configure(state='disabled')
        except Exception:
            say('计划已失效或尚未选择，请重新提出计划。')

    def confirm(approved):
        try:
            identifier = selected_plan()
            value = host.preview(identifier)
            if approved and not messagebox.askyesno('独立写入确认', '确认按本机预览写入新副本？源文件保持，执行备份、重开校验、渲染闸门和审计。\n\n' + value['before'][:700] + '\n→\n' + value['after'][:700], parent=root):
                return
            result = host.confirm_local(identifier, approved)
            say('已取消，未写入。' if result.get('status') == 'cancelled' else '本机事务已返回，请核对副本：' + str(result.get('output', '')))
        except Exception:
            say('操作未能确认完成；请核对本机备份和审计，不自动重试。')

    def rollback():
        try:
            identifier = selected_plan()
            if identifier not in host.history:
                raise ValueError()
            if messagebox.askyesno('独立撤销确认', '确认撤销本窗口此次已保存副本？保留源文件、备份和审计。', parent=root):
                host.rollback_local(identifier, True)
                say('已完成本机撤销。')
        except Exception:
            say('未取得有效撤销结果；请核对本机副本和审计。')

    actions = ttk.Frame(frame)
    actions.pack(fill='x', pady=8)
    ttk.Button(actions, text='本机确认写副本', command=lambda: confirm(True)).pack(side='left')
    ttk.Button(actions, text='取消计划', command=lambda: confirm(False)).pack(side='left', padx=8)
    ttk.Button(actions, text='独立撤销副本', command=rollback).pack(side='left')
    plans.bind('<<ListboxSelect>>', show)

    def poll():
        nonlocal visible
        if done.is_set():
            close()
            return
        with host.lock:
            try:
                if host.path is not None:
                    host._guard()
                    host._prune()
            except Exception:
                host.revoke()
                say('许可已过期、文件改变或账户失效，请重新登录并授权。')
            current = list(host.pending) + list(host.history)
            if current != visible:
                visible = current
                plans.delete(0, 'end')
                for identifier in visible:
                    plans.insert('end', ('已保存 ' if identifier in host.history else '待审阅 ') + identifier[:12])
                preview.configure(state='normal')
                preview.delete('1.0', 'end')
                preview.configure(state='disabled')
        root.after(250, poll)

    def close():
        host.logout()
        password.delete(0, 'end')
        root.destroy()

    root.protocol('WM_DELETE_WINDOW', close)
    root.after(250, poll)
    return root


def main():
    parser = argparse.ArgumentParser(description='LiX MCP stdio；本机登录及文件许可独立于其他窗口，默认不返回正文。')
    parser.add_argument('--headless', action='store_true', help='无控制窗口，仅协议诊断；无登录/文件授权入口，业务调用始终拒绝')
    args = parser.parse_args()
    logging.disable(logging.CRITICAL)
    host = McpHost()
    if args.headless:
        try:
            asyncio.run(serve(host))
        except Exception:
            sys.stderr.write('MCP transport stopped.\n')
        finally:
            host.logout()
        return
    done = threading.Event()
    try:
        root = control_window(host, done)
    except Exception:
        sys.stderr.write('MCP local authorization window unavailable.\n')
        return

    def protocol():
        try:
            asyncio.run(serve(host))
        except Exception:
            pass
        finally:
            done.set()

    thread = threading.Thread(target=protocol, name='office-agent-mcp-stdio', daemon=True)
    thread.start()
    try:
        root.mainloop()
    finally:
        host.logout()


if __name__ == '__main__':
    main()
