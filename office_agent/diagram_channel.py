"""Bounded localhost-only in-memory renderer wire format (not HTTP/API)."""
import json
import socket
import struct
import time

MAX_MESSAGE = 2_100_000  # <=1 MiB SVG plus worst-case JSON quoting and fixed keys.
MAGIC = b'OD1'


def read_exact(connection, count, deadline, stop=None):
    data = bytearray()
    while len(data) < count:
        if time.monotonic() >= deadline or (stop is not None and stop.is_set()):
            raise TimeoutError('renderer channel stopped')
        try:
            chunk = connection.recv(min(65536, count - len(data)))
        except socket.timeout:
            continue
        if not chunk:
            raise ValueError('incomplete renderer message')
        data.extend(chunk)
    if time.monotonic() >= deadline or stop is not None and stop.is_set():
        raise TimeoutError('diagram channel deadline')
    return bytes(data)


def send_message(connection, payload):
    data = json.dumps(payload, ensure_ascii=False, allow_nan=False,
                      separators=(',', ':')).encode('utf-8')
    if not 0 < len(data) <= MAX_MESSAGE:
        raise ValueError('renderer message too large')
    connection.sendall(struct.pack('!I', len(data)) + data)


def receive_message(connection, deadline, stop=None):
    size = struct.unpack('!I', read_exact(connection, 4, deadline, stop))[0]
    if not 0 < size <= MAX_MESSAGE:
        raise ValueError('renderer message too large')
    def unique(pairs):
        result = {}
        for key, value in pairs:
            if key in result:
                raise ValueError('duplicate renderer field')
            result[key] = value
        return result
    payload = json.loads(read_exact(connection, size, deadline, stop).decode('utf-8'),
                         object_pairs_hook=unique,
                         parse_constant=lambda _: (_ for _ in ()).throw(ValueError('invalid number')))
    if type(payload) is not dict:
        raise ValueError('renderer message is not an object')
    return payload
