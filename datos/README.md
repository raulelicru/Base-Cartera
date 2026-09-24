# Datos incluidos

## `catalogo_cp_sepomex.txt.gz`

Catálogo Nacional de Códigos Postales de México (SEPOMEX / Correos de México), reducido a una fila por
código postal: `Cp|Municipio|Estado|Zona|TipoAsentamiento` (32,467 CP).

- Fuente: datos públicos de SEPOMEX, tomados del paquete npm
  [`@webrek/mx-cp`](https://github.com/webrek/mx-cp) v0.2.0 (licencia MIT), que los distribuye ya procesados.
- La app lo usa automáticamente cuando no se ha cargado un catálogo propio. Si se sube el catálogo oficial
  más reciente desde la barra lateral, ese tiene prioridad.
- Se recomienda actualizarlo al menos una vez al año descargando el catálogo oficial de correosdemexico.gob.mx.
