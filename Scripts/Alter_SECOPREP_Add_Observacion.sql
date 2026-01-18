-- Script para agregar el campo Observacion a la tabla SECOPREP
-- Este campo almacenará las observaciones migradas desde el archivo OBSOPREP.DAT

ALTER TABLE [dbo].[SECOPREP]
ADD [Observacion] [nvarchar](512) NULL;
GO
