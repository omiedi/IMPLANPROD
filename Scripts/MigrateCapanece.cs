using System;
using System.IO;
using System.Text;
using System.Data.SqlClient;
using System.Collections.Generic;

namespace CapaneceMigration
{
    /// <summary>
    /// Utilidad para migrar datos desde CAPANECE.DAT (VB6) hacia SQL Server
    /// Basado en la estructura del archivo CAPANECE.RFS
    /// </summary>
    class Program
    {
        // Estructura del registro según CAPANECE.RFS
        const int RECORD_SIZE = 85; // Tamaño aproximado del registro
        
        // Posiciones y tamaños de campos según el .RFS
        const int POS_EQUIPO = 0;
        const int LEN_EQUIPO = 6;
        
        const int POS_COD_PIEZA = 6;
        const int LEN_COD_PIEZA = 2;
        
        const int POS_NUM_OPERACION = 8;
        const int LEN_NUM_OPERACION = 2;
        
        const int POS_DESCRIPCION = 10;
        const int LEN_DESCRIPCION = 48;
        
        const int POS_CANT_PROG = 58;
        const int LEN_CANT_PROG = 4; // Single (float)
        
        const int POS_CANT_TOTAL = 62;
        const int LEN_CANT_TOTAL = 4; // Single (float)
        
        const int POS_COD_PROD_DEST = 66;
        const int LEN_COD_PROD_DEST = 2;
        
        const int POS_HORAS_NEC = 68;
        const int LEN_HORAS_NEC = 4; // Single (float)
        
        const int POS_TOTAL_HORAS_PIEZA = 72;
        const int LEN_TOTAL_HORAS_PIEZA = 4; // Single (float)
        
        const int POS_TOTAL_HORAS_EQUIPO = 76;
        const int LEN_TOTAL_HORAS_EQUIPO = 4; // Single (float)

        static void Main(string[] args)
        {
            Console.WriteLine("=== Migración de CAPANECE.DAT a SQL Server ===");
            Console.WriteLine();

            // Rutas de archivos
            string datFilePath = @"c:\Users\Omar\source\repos\omardemaio\IMPLANPROD\Legacy_VB6\ArchivosDat\CAPANECE.DAT";
            
            // Verificar si existe el archivo
            if (!File.Exists(datFilePath))
            {
                Console.WriteLine($"ERROR: No se encontró el archivo {datFilePath}");
                return;
            }

            // IMPORTANTE: Configura tu cadena de conexión aquí
            string connectionString = "CONFIGURA_TU_CADENA_DE_CONEXION_AQUI";
            
            Console.WriteLine("NOTA: Antes de ejecutar, configura la cadena de conexión en el código.");
            Console.WriteLine();
            Console.WriteLine("Ejemplo de cadena de conexión:");
            Console.WriteLine("Server=localhost;Database=TuBaseDeDatos;Trusted_Connection=True;");
            Console.WriteLine();
            Console.WriteLine("Presiona ENTER para continuar (asegúrate de configurar la conexión primero)...");
            Console.ReadLine();

            try
            {
                // Leer archivo .DAT
                byte[] fileData = File.ReadAllBytes(datFilePath);
                Console.WriteLine($"Archivo leído: {fileData.Length} bytes");
                
                // Calcular número de registros
                int recordCount = fileData.Length / RECORD_SIZE;
                Console.WriteLine($"Registros estimados: {recordCount}");
                Console.WriteLine();

                List<CapaneceRecord> records = new List<CapaneceRecord>();

                // Procesar cada registro
                for (int i = 0; i < recordCount; i++)
                {
                    int offset = i * RECORD_SIZE;
                    
                    if (offset + RECORD_SIZE > fileData.Length)
                        break;

                    try
                    {
                        CapaneceRecord record = ParseRecord(fileData, offset);
                        
                        // Solo agregar si tiene datos válidos
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

                // Mostrar algunos registros de ejemplo
                Console.WriteLine("=== PRIMEROS 5 REGISTROS ===");
                for (int i = 0; i < Math.Min(5, records.Count); i++)
                {
                    var r = records[i];
                    Console.WriteLine($"Registro {i + 1}:");
                    Console.WriteLine($"  Equipo: {r.Equipo}");
                    Console.WriteLine($"  Horas Necesarias: {r.HorasNecesarias}");
                    Console.WriteLine($"  Total Horas Equipo: {r.TotalHorasPorEquipo}");
                    Console.WriteLine();
                }

                // Generar script SQL de inserción
                Console.WriteLine("=== GENERANDO SCRIPT SQL ===");
                GenerateSqlInsertScript(records, @"c:\Users\Omar\source\repos\omardemaio\IMPLANPROD\Scripts\Insert_CAPANECE_Migracion.sql");
                Console.WriteLine("Script SQL generado en: Scripts\\Insert_CAPANECE_Migracion.sql");
                Console.WriteLine();
                Console.WriteLine("NOTA: Revisa el script SQL antes de ejecutarlo en la base de datos.");

                // Si deseas insertar directamente (descomentar y configurar la conexión)
                /*
                Console.WriteLine("¿Deseas insertar los datos directamente en la BD? (S/N)");
                string response = Console.ReadLine();
                if (response?.ToUpper() == "S")
                {
                    InsertRecordsToDatabase(records, connectionString);
                }
                */
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
            var record = new CapaneceRecord();

            // Equipo (6 caracteres)
            record.Equipo = ReadString(data, offset + POS_EQUIPO, LEN_EQUIPO);

            // Código de pieza (entero corto)
            record.CodigoPieza = ReadInt16(data, offset + POS_COD_PIEZA);

            // Número de operación
            record.NumeroOperacion = ReadInt16(data, offset + POS_NUM_OPERACION);

            // Descripción
            record.Descripcion = ReadString(data, offset + POS_DESCRIPCION, LEN_DESCRIPCION);

            // Cantidad programada (Single)
            record.CantidadProgramada = ReadSingle(data, offset + POS_CANT_PROG);

            // Cantidad total
            record.CantidadTotalPieza = ReadSingle(data, offset + POS_CANT_TOTAL);

            // Código producto destino
            record.CodigoProductoDestino = ReadInt16(data, offset + POS_COD_PROD_DEST);

            // Horas necesarias
            record.HorasNecesarias = ReadSingle(data, offset + POS_HORAS_NEC);

            // Total horas por pieza
            record.TotalHorasPorPieza = ReadSingle(data, offset + POS_TOTAL_HORAS_PIEZA);

            // Total horas por equipo
            record.TotalHorasPorEquipo = ReadSingle(data, offset + POS_TOTAL_HORAS_EQUIPO);

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
                writer.WriteLine($"-- Generado: {DateTime.Now}");
                writer.WriteLine($"-- Total de registros: {records.Count}");
                writer.WriteLine();
                writer.WriteLine("-- IMPORTANTE: Revisa este script antes de ejecutarlo");
                writer.WriteLine("-- Asegúrate de que la tabla CAPANECE ya existe");
                writer.WriteLine();
                writer.WriteLine("SET NOCOUNT ON;");
                writer.WriteLine("BEGIN TRANSACTION;");
                writer.WriteLine();

                int batchSize = 100;
                for (int i = 0; i < records.Count; i++)
                {
                    var r = records[i];

                    writer.WriteLine($"INSERT INTO CAPANECE (Equipo, CodigoPieza, NumeroOperacion, Descripcion, CantidadProgramada, CantidadTotalPieza, CodigoProductoDestino, HorasNecesarias, TotalHorasPorPieza, TotalHorasPorEquipo)");
                    writer.WriteLine($"VALUES ('{EscapeSql(r.Equipo)}', {(r.CodigoPieza == 0 ? "NULL" : r.CodigoPieza.ToString())}, {(r.NumeroOperacion == 0 ? "NULL" : r.NumeroOperacion.ToString())}, '{EscapeSql(r.Descripcion)}', {FormatDecimal(r.CantidadProgramada)}, {FormatDecimal(r.CantidadTotalPieza)}, {(r.CodigoProductoDestino == 0 ? "NULL" : r.CodigoProductoDestino.ToString())}, {FormatDecimal(r.HorasNecesarias)}, {FormatDecimal(r.TotalHorasPorPieza)}, {FormatDecimal(r.TotalHorasPorEquipo)});");
                    
                    // Crear lotes para mejor rendimiento
                    if ((i + 1) % batchSize == 0 && i < records.Count - 1)
                    {
                        writer.WriteLine();
                        writer.WriteLine("-- Lote completado");
                        writer.WriteLine();
                    }
                }

                writer.WriteLine();
                writer.WriteLine("COMMIT TRANSACTION;");
                writer.WriteLine($"PRINT 'Migración completada: {records.Count} registros insertados';");
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
            if (value == 0)
                return "NULL";
            
            return value.ToString("0.0000").Replace(",", ".");
        }

        static void InsertRecordsToDatabase(List<CapaneceRecord> records, string connectionString)
        {
            Console.WriteLine("Insertando registros en la base de datos...");
            
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                
                int inserted = 0;
                foreach (var record in records)
                {
                    try
                    {
                        using (SqlCommand cmd = new SqlCommand(@"
                            INSERT INTO CAPANECE (Equipo, CodigoPieza, NumeroOperacion, Descripcion, 
                                CantidadProgramada, CantidadTotalPieza, CodigoProductoDestino, 
                                HorasNecesarias, TotalHorasPorPieza, TotalHorasPorEquipo)
                            VALUES (@Equipo, @CodigoPieza, @NumeroOperacion, @Descripcion, 
                                @CantidadProgramada, @CantidadTotalPieza, @CodigoProductoDestino, 
                                @HorasNecesarias, @TotalHorasPorPieza, @TotalHorasPorEquipo)", conn))
                        {
                            cmd.Parameters.AddWithValue("@Equipo", record.Equipo);
                            cmd.Parameters.AddWithValue("@CodigoPieza", record.CodigoPieza == 0 ? (object)DBNull.Value : record.CodigoPieza);
                            cmd.Parameters.AddWithValue("@NumeroOperacion", record.NumeroOperacion == 0 ? (object)DBNull.Value : record.NumeroOperacion);
                            cmd.Parameters.AddWithValue("@Descripcion", string.IsNullOrEmpty(record.Descripcion) ? (object)DBNull.Value : record.Descripcion);
                            cmd.Parameters.AddWithValue("@CantidadProgramada", record.CantidadProgramada == 0 ? (object)DBNull.Value : (decimal)record.CantidadProgramada);
                            cmd.Parameters.AddWithValue("@CantidadTotalPieza", record.CantidadTotalPieza == 0 ? (object)DBNull.Value : (decimal)record.CantidadTotalPieza);
                            cmd.Parameters.AddWithValue("@CodigoProductoDestino", record.CodigoProductoDestino == 0 ? (object)DBNull.Value : record.CodigoProductoDestino);
                            cmd.Parameters.AddWithValue("@HorasNecesarias", record.HorasNecesarias == 0 ? (object)DBNull.Value : (decimal)record.HorasNecesarias);
                            cmd.Parameters.AddWithValue("@TotalHorasPorPieza", record.TotalHorasPorPieza == 0 ? (object)DBNull.Value : (decimal)record.TotalHorasPorPieza);
                            cmd.Parameters.AddWithValue("@TotalHorasPorEquipo", record.TotalHorasPorEquipo == 0 ? (object)DBNull.Value : (decimal)record.TotalHorasPorEquipo);

                            cmd.ExecuteNonQuery();
                            inserted++;
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error al insertar registro: {ex.Message}");
                    }
                }
                
                Console.WriteLine($"Registros insertados: {inserted}");
            }
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
