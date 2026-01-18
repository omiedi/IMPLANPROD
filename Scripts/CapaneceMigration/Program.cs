using System;
using System.IO;
using System.Text;

namespace CapaneceMigration
{
    /// <summary>
    /// Utilidad para migrar datos desde CAPANECE.DAT (VB6) hacia SQL Server
    /// Basado en la estructura del archivo CAPANECE.RFS
    /// </summary>
    class Program
    {
        // Estructura del registro según CAPANECE.RFS (85 bytes por registro)
        const int RECORD_SIZE = 85;
        
        // Posiciones y tamaños de campos
        const int POS_EQUIPO = 0;
        const int LEN_EQUIPO = 6;
        
        const int POS_COD_PIEZA = 6;
        const int POS_NUM_OPERACION = 8;
        const int POS_DESCRIPCION = 10;
        const int LEN_DESCRIPCION = 48;
        const int POS_CANT_PROG = 58;
        const int POS_CANT_TOTAL = 62;
        const int POS_COD_PROD_DEST = 66;
        const int POS_HORAS_NEC = 68;
        const int POS_TOTAL_HORAS_PIEZA = 72;
        const int POS_TOTAL_HORAS_EQUIPO = 76;

        static void Main(string[] args)
        {
            Console.WriteLine("=== Migración de CAPANECE.DAT a SQL Server ===");
            Console.WriteLine();

            string datFilePath = @"..\..\Legacy_VB6\ArchivosDat\CAPANECE.DAT";
            
            if (!File.Exists(datFilePath))
            {
                Console.WriteLine($"ERROR: No se encontró el archivo {datFilePath}");
                Console.WriteLine("Presiona ENTER para salir...");
                Console.ReadLine();
                return;
            }

            try
            {
                byte[] fileData = File.ReadAllBytes(datFilePath);
                Console.WriteLine($"Archivo leído: {fileData.Length:N0} bytes");
                
                int recordCount = fileData.Length / RECORD_SIZE;
                Console.WriteLine($"Registros estimados: {recordCount}");
                Console.WriteLine();

                List<CapaneceRecord> records = new List<CapaneceRecord>();

                for (int i = 0; i < recordCount; i++)
                {
                    int offset = i * RECORD_SIZE;
                    
                    if (offset + RECORD_SIZE > fileData.Length)
                        break;

                    try
                    {
                        CapaneceRecord record = ParseRecord(fileData, offset);
                        
                        if (!string.IsNullOrWhiteSpace(record.Equipo))
                        {
                            records.Add(record);
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error al procesar registro {i}: {ex.Message}");
                    }
                }

                Console.WriteLine($"Registros válidos encontrados: {records.Count}");
                Console.WriteLine();

                // Mostrar primeros 10 registros
                Console.WriteLine("=== PRIMEROS 10 REGISTROS ===");
                for (int i = 0; i < Math.Min(10, records.Count); i++)
                {
                    var r = records[i];
                    Console.WriteLine($"#{i + 1}: Equipo={r.Equipo}, HorasNec={r.HorasNecesarias:F4}, TotalHsEq={r.TotalHorasPorEquipo:F4}");
                }
                Console.WriteLine();

                // Generar script SQL
                string outputScript = @"..\Insert_CAPANECE_Migracion.sql";
                GenerateSqlInsertScript(records, outputScript);
                Console.WriteLine($"✓ Script SQL generado: {Path.GetFullPath(outputScript)}");
                Console.WriteLine();
                Console.WriteLine("IMPORTANTE: Revisa el script SQL antes de ejecutarlo en la base de datos.");
                Console.WriteLine($"Total de registros en el script: {records.Count}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"ERROR: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }

            Console.WriteLine();
            Console.WriteLine("Presiona ENTER para salir...");
            Console.ReadLine();
        }

        static CapaneceRecord ParseRecord(byte[] data, int offset)
        {
            var record = new CapaneceRecord
            {
                Equipo = ReadString(data, offset + POS_EQUIPO, LEN_EQUIPO),
                CodigoPieza = ReadInt16(data, offset + POS_COD_PIEZA),
                NumeroOperacion = ReadInt16(data, offset + POS_NUM_OPERACION),
                Descripcion = ReadString(data, offset + POS_DESCRIPCION, LEN_DESCRIPCION),
                CantidadProgramada = ReadSingle(data, offset + POS_CANT_PROG),
                CantidadTotalPieza = ReadSingle(data, offset + POS_CANT_TOTAL),
                CodigoProductoDestino = ReadInt16(data, offset + POS_COD_PROD_DEST),
                HorasNecesarias = ReadSingle(data, offset + POS_HORAS_NEC),
                TotalHorasPorPieza = ReadSingle(data, offset + POS_TOTAL_HORAS_PIEZA),
                TotalHorasPorEquipo = ReadSingle(data, offset + POS_TOTAL_HORAS_EQUIPO)
            };

            return record;
        }

        static string ReadString(byte[] data, int offset, int length)
        {
            if (offset + length > data.Length)
                return string.Empty;

            return Encoding.ASCII.GetString(data, offset, length).Trim('\0', ' ');
        }

        static short ReadInt16(byte[] data, int offset)
        {
            if (offset + 2 > data.Length)
                return 0;

            return BitConverter.ToInt16(data, offset);
        }

        static float ReadSingle(byte[] data, int offset)
        {
            if (offset + 4 > data.Length)
                return 0f;

            return BitConverter.ToSingle(data, offset);
        }

        static void GenerateSqlInsertScript(List<CapaneceRecord> records, string outputPath)
        {
            using (StreamWriter writer = new StreamWriter(outputPath, false, Encoding.UTF8))
            {
                writer.WriteLine("-- Script de migración de datos CAPANECE.DAT");
                writer.WriteLine($"-- Generado: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
                writer.WriteLine($"-- Total de registros: {records.Count}");
                writer.WriteLine();
                writer.WriteLine("-- IMPORTANTE: Revisa este script antes de ejecutarlo");
                writer.WriteLine("-- Asegúrate de que la tabla CAPANECE ya existe en la base de datos");
                writer.WriteLine();
                writer.WriteLine("USE [TU_BASE_DE_DATOS]; -- Cambia esto por tu base de datos");
                writer.WriteLine("GO");
                writer.WriteLine();
                writer.WriteLine("SET NOCOUNT ON;");
                writer.WriteLine("BEGIN TRANSACTION;");
                writer.WriteLine();
                writer.WriteLine("-- Limpiar datos existentes (opcional, comenta si no deseas borrar)");
                writer.WriteLine("-- DELETE FROM CAPANECE;");
                writer.WriteLine();

                int batchSize = 500;
                int batchCount = 0;
                
                for (int i = 0; i < records.Count; i++)
                {
                    var r = records[i];

                    writer.Write($"INSERT INTO CAPANECE (Equipo, CodigoPieza, NumeroOperacion, Descripcion, ");
                    writer.Write($"CantidadProgramada, CantidadTotalPieza, CodigoProductoDestino, ");
                    writer.WriteLine($"HorasNecesarias, TotalHorasPorPieza, TotalHorasPorEquipo) VALUES (");
                    writer.Write($"    '{EscapeSql(r.Equipo)}', ");
                    writer.Write($"{NullOrValue(r.CodigoPieza)}, ");
                    writer.Write($"{NullOrValue(r.NumeroOperacion)}, ");
                    writer.Write($"'{EscapeSql(r.Descripcion)}', ");
                    writer.Write($"{FormatDecimal(r.CantidadProgramada)}, ");
                    writer.Write($"{FormatDecimal(r.CantidadTotalPieza)}, ");
                    writer.Write($"{NullOrValue(r.CodigoProductoDestino)}, ");
                    writer.Write($"{FormatDecimal(r.HorasNecesarias)}, ");
                    writer.Write($"{FormatDecimal(r.TotalHorasPorPieza)}, ");
                    writer.WriteLine($"{FormatDecimal(r.TotalHorasPorEquipo)});");
                    
                    batchCount++;
                    
                    // Crear lotes para mejor rendimiento
                    if (batchCount >= batchSize && i < records.Count - 1)
                    {
                        writer.WriteLine();
                        writer.WriteLine($"-- Lote completado: {i + 1} registros");
                        writer.WriteLine();
                        batchCount = 0;
                    }
                }

                writer.WriteLine();
                writer.WriteLine("COMMIT TRANSACTION;");
                writer.WriteLine("GO");
                writer.WriteLine();
                writer.WriteLine($"PRINT 'Migración completada: {records.Count} registros insertados';");
                writer.WriteLine("GO");
                writer.WriteLine();
                writer.WriteLine("-- Verificar registros insertados");
                writer.WriteLine("SELECT COUNT(*) AS TotalRegistros FROM CAPANECE;");
                writer.WriteLine("SELECT TOP 10 * FROM CAPANECE ORDER BY Id DESC;");
            }
        }

        static string EscapeSql(string value)
        {
            if (string.IsNullOrEmpty(value))
                return string.Empty;
            
            return value.Replace("'", "''");
        }

        static string FormatDecimal(float value)
        {
            if (value == 0 || float.IsNaN(value) || float.IsInfinity(value))
                return "NULL";
            
            return value.ToString("0.0000", System.Globalization.CultureInfo.InvariantCulture);
        }

        static string NullOrValue(short value)
        {
            if (value == 0)
                return "NULL";
            
            return value.ToString();
        }
    }

    class CapaneceRecord
    {
        public string Equipo { get; set; } = string.Empty;
        public short CodigoPieza { get; set; }
        public short NumeroOperacion { get; set; }
        public string Descripcion { get; set; } = string.Empty;
        public float CantidadProgramada { get; set; }
        public float CantidadTotalPieza { get; set; }
        public short CodigoProductoDestino { get; set; }
        public float HorasNecesarias { get; set; }
        public float TotalHorasPorPieza { get; set; }
        public float TotalHorasPorEquipo { get; set; }
    }
}
