-- Fila de cada cuenta en la hoja original y ruta del archivo original en Storage,
-- para poder volver a descargar la cartera tal como se envió con las columnas anexadas.
alter table public.base_gestion add column fila_excel integer;
alter table public.corridas add column archivo_original text;
