# -*- coding: utf-8 -*-
"""
企业微信内部报价助手 - 回调服务
在企业微信群 @机器人 发询价 → 企微回调本服务 → 报价引擎计算 → 机器人回复报价

部署：
  pip install -r requirements.txt
  export WECOM_TOKEN=xxx WECOM_AESKEY=xxx WECOM_CORPID=xxx
  uvicorn wecom_bot:app --host 0.0.0.0 --port 8000
"""
import os
import re
import base64
import hashlib
import struct
import socket
import time
import xml.etree.ElementTree as ET
from flask import Flask, request, Response
from quote_engine import quote_us_air, format_reply

app = Flask(__name__)

TOKEN = os.environ.get("WECOM_TOKEN", "")
AES_KEY = os.environ.get("WECOM_AESKEY", "")   # 43位 EncodingAESKey
CORPID = os.environ.get("WECOM_CORPID", "")


# ---------- 企微消息加解密（标准 WXBizMsgCrypt 算法） ----------

def _sign(token, timestamp, nonce, encrypt):
    return hashlib.sha1("".join(sorted([token, timestamp, nonce, encrypt])).encode()).hexdigest()


def decrypt_message(encrypt_b64: str) -> str:
    key = base64.b64decode(AES_KEY + "=")
    cipher = AES.new(key, AES.MODE_CBC, key[:16])
    plain = cipher.decrypt(base64.b64decode(encrypt_b64))
    pad = plain[-1]
    content = plain[16:-pad]
    msg_len = struct.unpack(">I", content[:4])[0]
    corp_id = content[4 + msg_len:].decode()
    if corp_id != CORPID:
        raise ValueError("corpid 校验失败")
    return content[4:4 + msg_len].decode()


def encrypt_message(reply_xml: str, nonce: str) -> str:
    import Crypto.Random
    key = base64.b64decode(AES_KEY + "=")
    content = (Crypto.Random.get_random_bytes(16) + struct.pack(">I", len(reply_xml.encode()))
               + reply_xml.encode() + CORPID.encode())
    pad = 32 - len(content) % 32
    content += bytes([pad]) * pad
    cipher = AES.new(key, AES.MODE_CBC, key[:16])
    return base64.b64encode(cipher.encrypt(content)).decode()


def wrap_reply(xml: str, nonce: str, timestamp=None):
    ts = timestamp or str(int(time.time()))
    encrypt = encrypt_message(xml, nonce)
    sig = _sign(TOKEN, ts, nonce, encrypt)
    return (f"<xml><Encrypt><![CDATA[{encrypt}]]></Encrypt>"
            f"<MsgSignature><![CDATA[{sig}]]></MsgSignature>"
            f"<TimeStamp>{ts}</TimeStamp><Nonce><![CDATA[{nonce}]]></Nonce></xml>")


# ---------- 路由 ----------

@app.route("/wecom/callback", methods=["GET"])
def verify():
    """URL 验证（企微后台配置回调时触发）"""
    args = request.args
    sig = _sign(TOKEN, args.get("timestamp", ""), args.get("nonce", ""), args.get("echostr", ""))
    if sig != args.get("msg_signature"):
        return "bad signature", 403
    return decrypt_message(args.get("echostr", "")), 200


@app.route("/wecom/callback", methods=["POST"])
def callback():
    args = request.args
    body = ET.fromstring(request.data)
    encrypt = body.findtext("Encrypt")
    sig = _sign(TOKEN, args.get("timestamp", ""), args.get("nonce", ""), encrypt)
    if sig != args.get("msg_signature"):
        return "bad signature", 403
    msg_xml = ET.fromstring(decrypt_message(encrypt))
    msg_type = msg_xml.findtext("MsgType")
    nonce = args.get("nonce", "")

    if msg_type == "text":
        text = msg_xml.findtext("Content") or ""
        result = quote_us_air(text)
        reply = format_reply(result)
    elif msg_type == "image":
        reply = "已收到图片。请补充文字信息：尺寸(长*宽*高cm)、单件重量、箱数、目的地邮编，例如：\n39*28*25cm 17kg/箱 2箱 76117"
    else:
        reply = "请发送询价内容，例如：量勺7件套 39*28*25cm 17kg/箱 2箱 76117"

    from_user = msg_xml.findtext("FromUserName") or ""
    to_user = msg_xml.findtext("ToUserName") or ""
    reply_xml = (f"<xml><ToUserName><![CDATA[{from_user}]]></ToUserName>"
                 f"<FromUserName><![CDATA[{to_user}]]></FromUserName>"
                 f"<CreateTime>{int(time.time())}</CreateTime>"
                 f"<MsgType><![CDATA[text]]></MsgType>"
                 f"<Content><![CDATA[{reply}]]></Content></xml>")
    return Response(wrap_reply(reply_xml, nonce), mimetype="application/xml")


@app.route("/health")
def health():
    return "ok"


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8000)
