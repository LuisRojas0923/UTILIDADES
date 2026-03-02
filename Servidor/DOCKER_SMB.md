# Ejecutar el servidor en Docker (Linux) y leer Excel desde 192.168.0.3

Sí se puede. El contenedor **no** puede usar rutas UNC de Windows (`\\192.168.0.3\...`). La solución es **montar el recurso compartido SMB en el host Linux** y pasar esa carpeta al contenedor con un volumen.

---

## 1. Resumen

1. En el **servidor Linux** montas las carpetas de red de `192.168.0.3` en un directorio local (p. ej. `/mnt/excel`).
2. Arrancas el contenedor con un **volumen** que mapea ese directorio a una ruta dentro del contenedor (p. ej. `/data/excel`).
3. Defines la variable de entorno **`EXCEL_BASE_PATH=/data/excel`** para que el script use esas rutas en lugar de `\\192.168.0.3\...`.

---

## 2. Montar el recurso compartido en el host Linux

En el servidor Linux (donde corre Docker) necesitas montar los recursos compartidos que usa el script:

- `\\192.168.0.3\Procesos Comunes SGI`
- `\\192.168.0.3\Postventa`
- `\\192.168.0.3\Control Presupuestal` (archivo MEMOFICHAS – consulta memofichas v2.xlsx)

### 2.1 Instalar cliente SMB/CIFS

```bash
# Debian/Ubuntu
sudo apt-get update
sudo apt-get install -y cifs-utils

# RHEL/CentOS/Rocky
sudo dnf install -y cifs-utils
```

### 2.2 Crear archivo de credenciales (recomendado)

Crea un archivo que solo root pueda leer (no lo subas a Git):

```bash
sudo mkdir -p /etc/smb-credentials
sudo nano /etc/smb-credentials/sgi
```

Contenido (ajusta usuario y contraseña del dominio/Windows):

```ini
username=usuario_dominio
password=contraseña
domain=DOMINIO
```

Permisos:

```bash
sudo chmod 600 /etc/smb-credentials/sgi
```

### 2.3 Puntos de montaje

Crea el directorio donde verás los Excel (la misma estructura que bajo `192.168.0.3`):

```bash
sudo mkdir -p /mnt/excel
sudo mkdir -p "/mnt/excel/Procesos Comunes SGI"
sudo mkdir -p "/mnt/excel/Postventa"
sudo mkdir -p "/mnt/excel/Control Presupuestal"
```

Monta cada recurso (los nombres con espacios van entre comillas):

```bash
# Recurso "Procesos Comunes SGI"
sudo mount -t cifs "//192.168.0.3/Procesos Comunes SGI" "/mnt/excel/Procesos Comunes SGI" \
  -o credentials=/etc/smb-credentials/sgi,uid=$(id -u),gid=$(id -g),file_mode=0644,dir_mode=0755

# Recurso "Postventa"
sudo mount -t cifs "//192.168.0.3/Postventa" "/mnt/excel/Postventa" \
  -o credentials=/etc/smb-credentials/sgi,uid=$(id -u),gid=$(id -g),file_mode=0644,dir_mode=0755

# Recurso "Control Presupuestal" (MEMOFICHAS)
sudo mount -t cifs "//192.168.0.3/Control Presupuestal" "/mnt/excel/Control Presupuestal" \
  -o credentials=/etc/smb-credentials/sgi,uid=$(id -u),gid=$(id -g),file_mode=0644,dir_mode=0755
```

Comprueba que se ven los archivos:

```bash
ls -la "/mnt/excel/Procesos Comunes SGI/Costos/INFORME DE ORDENES/"
ls -la "/mnt/excel/Postventa/MANTENIMIENTO Y SERVICIO POSTVENTA/- GESTION ORDENES DE SERVICIO/CENTRO LOGÍSTICO/"
ls -la "/mnt/excel/Control Presupuestal/"
```

### 2.4 Montaje automático al arrancar (opcional)

Para que los montajes se hagan al reiniciar el servidor, añade en `/etc/fstab`:

```bash
sudo nano /etc/fstab
```

Líneas (en una sola línea por entrada; aquí partidas por legibilidad):

```
//192.168.0.3/Procesos Comunes SGI  /mnt/excel/Procesos Comunes SGI  cifs  credentials=/etc/smb-credentials/sgi,uid=1000,gid=1000,file_mode=0644,dir_mode=0755,_netdev  0  0
//192.168.0.3/Postventa             /mnt/excel/Postventa             cifs  credentials=/etc/smb-credentials/sgi,uid=1000,gid=1000,file_mode=0644,dir_mode=0755,_netdev  0  0
//192.168.0.3/Control Presupuestal  /mnt/excel/Control Presupuestal   cifs  credentials=/etc/smb-credentials/sgi,uid=1000,gid=1000,file_mode=0644,dir_mode=0755,_netdev  0  0
```

Sustituye `uid=1000,gid=1000` por el usuario que ejecuta Docker (`id -u` y `id -g`). `_netdev` hace que espere a que la red esté lista.

---

## 3. Ejecutar el contenedor con el volumen

Con los recursos ya montados en `/mnt/excel`:

```bash
docker run -d \
  --name solid-etl \
  -p 8099:8099 \
  -e EXCEL_BASE_PATH=/data/excel \
  -v /mnt/excel:/data/excel:ro \
  tu-imagen-servidor
```

- **`EXCEL_BASE_PATH=/data/excel`**: el script usa rutas bajo `/data/excel` (equivalente a `\\192.168.0.3\...`).
- **`-v /mnt/excel:/data/excel:ro`**: el host montado se ve dentro del contenedor en `/data/excel`. `:ro` es opcional (solo lectura).

Así el script en Docker leerá los mismos Excel que en Windows, pero por la carpeta montada.

---

## 4. Docker Compose (ejemplo)

```yaml
services:
  servidor-sync:
    build: .
    ports:
      - "8099:8099"
    environment:
      - EXCEL_BASE_PATH=/data/excel
    volumes:
      - /mnt/excel:/data/excel:ro
    restart: unless-stopped
```

Requisito: en el host, `/mnt/excel` debe estar montado como en la sección 2 **antes** de levantar el compose.

---

## 5. Cambios en el script

El script `upload_buffer_polars.py` ya soporta este esquema:

- Si existe la variable de entorno **`EXCEL_BASE_PATH`**, construye las rutas de los tres Excel bajo esa carpeta (Linux/Docker).
- Si no existe, usa las rutas UNC de Windows (`\\192.168.0.3\...`).

No hace falta tocar código al pasar de Windows a Docker; solo configurar montaje y variable de entorno.

---

## 6. Comprobar conectividad

Desde el host Linux (o desde un contenedor con `curl`):

- **Red:** `ping 192.168.0.3`
- **Puerto SMB (445):** `nc -zv 192.168.0.3 445`
- **PostgreSQL (desde el contenedor):** que el contenedor pueda alcanzar `192.168.0.21:5432` (misma red o reglas de firewall que en Windows).

Si el ERP invoca al servicio por IP/hostname del servidor Linux, el puerto 8099 debe estar abierto en el firewall hacia ese cliente.
