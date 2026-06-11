<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Admin — Rendiciones</title>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600&display=swap" rel="stylesheet">
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Inter',sans-serif;background:#F0F4F8;min-height:100vh;display:flex;align-items:center;justify-content:center}
.card{background:#fff;border-radius:16px;padding:40px 36px;width:100%;max-width:380px;box-shadow:0 4px 24px rgba(0,0,0,.07)}
h1{font-size:20px;font-weight:600;color:#1F4E79;margin-bottom:6px}
p{font-size:14px;color:#6B7A8D;margin-bottom:24px}
label{font-size:13px;font-weight:500;color:#3D4F63;display:block;margin-bottom:6px}
input{width:100%;border:1.5px solid #D1DBE8;border-radius:10px;padding:11px 14px;font-size:14px;font-family:inherit;outline:none}
input:focus{border-color:#1F4E79}
.error{font-size:13px;color:#9B1C1C;background:#FDEAEA;padding:10px 14px;border-radius:8px;margin-bottom:16px}
button{width:100%;background:#1F4E79;color:#fff;border:none;border-radius:10px;padding:12px;font-size:15px;font-weight:500;cursor:pointer;margin-top:18px}
button:hover{background:#163D61}
</style>
</head>
<body>
<div class="card">
  <h1>🔐 Panel Admin</h1>
  <p>Acceso exclusivo para administradores</p>
  {% if error %}<div class="error">{{ error }}</div>{% endif %}
  <form method="POST">
    <label>Contraseña</label>
    <input type="password" name="password" autofocus placeholder="••••••••">
    <button type="submit">Entrar</button>
  </form>
</div>
</body>
</html>
