using IMPLANPROD.Shared.Entities;
using Microsoft.EntityFrameworkCore;
using Shared.Entities;


namespace IMPLANPROD.Server.Data
{
    /// <summary>
    /// Contexto de base de datos principal de la aplicación.
    /// Maneja la conexión y el mapeo de entidades con la base de datos.
    /// </summary>
    public class DataContext : DbContext
    {
        /// <summary>
        /// Constructor que recibe las opciones de configuración del contexto
        /// </summary>
        /// <param name="options">Opciones de configuración de Entity Framework</param>
        public DataContext(DbContextOptions<DataContext> options) : base(options)
        { }

        #region DbSets
        /// <summary>
        /// Tabla de EmailConfiguration
        /// </summary>
        public required DbSet<EmailConfiguration> EmailConfiguration { get; set; }

        /// <summary>
        /// Tabla de Estados
        /// </summary>
        public required DbSet<State> States { get; set; }

        /// <summary>
        /// Tabla de Ciudades
        /// </summary>
        public required DbSet<City> Cities { get; set; }

        /// <summary>
        /// Tabla de Países
        /// </summary>
        public required DbSet<Country> Countries { get; set; }

        /// <summary>
        /// Tabla de Dependencias
        /// </summary>
        public required DbSet<Dependencia> Dependencias { get; set; }

        /// <summary>
        /// Tabla de Empresas
        /// </summary>
        public required DbSet<Empresa> Empresas { get; set; }
        /// <summary>
        /// Tabla de Sectores
        /// </summary>
        public required DbSet<Sector> Sectores { get; set; }

        /// <summary>
        /// Tabla de Permisos
        /// </summary>
        public required DbSet<Permiso> Permisos { get; set; }

        /// <summary>
        /// Tabla de Perfiles
        /// </summary>
        public required DbSet<Perfil> Perfiles { get; set; }

        /// <summary>
        /// Tabla de Usuarios
        /// </summary>
        public required DbSet<Usuario> Usuarios { get; set; }

        /// <summary>
        /// Tabla de relaciones Usuario-Sector Productivo
        /// </summary>
        public required DbSet<UsuariosSectorProd> UsuariosSectorProds { get; set; }

        /// <summary>
        /// Tabla de Unidades de Medida
        /// </summary>
        public required DbSet<UnidadMedidas> UnidadMedidas { get; set; }

        /// <summary>
        /// Tabla de Deudores
        /// </summary>
        public required DbSet<Deudor12> Clientes { get; set; }

        /// <summary>
        /// Tabla de Deudores
        /// </summary>
        public required DbSet<Proved12> Proveedores { get; set; }

        /// <summary>
        /// Tabla de Reportes Personalizados
        /// </summary>
        public required DbSet<ReportePersonalizado> ReportesPersonalizados { get; set; }

        /// <summary>
        /// Tabla de Tadeposis
        /// </summary>
        public DbSet<Tadeposi> Tadeposi { get; set; }

        /// <summary>
        /// Tabla de Tatimat
        /// </summary>
        public DbSet<Tatimat> Tatimat { get; set; }

        /// <summary>
        /// Tabla de Familias
        /// </summary>
        public DbSet<Familia> Familias { get; set; }
        /// <summary>
        /// Tabla de Tamoneds
        /// </summary>
        public DbSet<Tamoned> Tamoneds { get; set; }

        /// <summary>
        /// Tabla de Tamoneds
        /// </summary>
        public DbSet<Grupiva> Grupivas { get; set; }

        /// <summary>
        /// Tabla de Tamoneds
        /// </summary>
        public DbSet<Catoper> Catopers { get; set; }
        /// <summary>
        /// Tabla de Tamoneds
        /// </summary>
        public DbSet<SecProd> SecProds { get; set; }

        /// <summary>
        /// Tabla padmaq
        /// </summary>
        public DbSet<PADMAQ> padmaqs { get; set; }

        /// <summary>
        /// Tabla Master
        /// </summary>
        public DbSet<Master> masters { get; set; }

        /// <summary>
        /// Tabla Formulas
        /// </summary>
        public DbSet<Formulas> Formulas { get; set; }

        /// <summary>
        /// Tabla Formulas
        /// </summary>
        public DbSet<Secoper> Secopers { get; set; }

        /// <summary>
        /// Tabla UploadResult
        /// </summary>
        public DbSet<UploadResult> UploadResult => Set<UploadResult>();


        /// <summary>
        /// Tabla UploadResult
        /// </summary>
        public DbSet<UploadMaster> UploadMaster { get; set; }

        /// <summary>
        /// Tabla Movisto
        /// </summary>
        public DbSet<Movisto> Movistos { get; set; }

        /// <summary>
        /// Tabla Movirepa (Movimientos de Reparación)
        /// </summary>
        public DbSet<Movirepa> Movirepas { get; set; }

        /// <summary>
        /// Tabla Ordefabr
        /// </summary>
        public DbSet<OrdeFabr> Ordefabrs { get; set; }

        /// <summary>
        /// Tabla Orfapen
        /// </summary>
        public DbSet<Orfapen> Orfapens { get; set; }

        /// <summary>
        /// Tabla Reserva
        /// </summary>
        public DbSet<Reserva> Reservas { get; set; }
        /// <summary>
        /// Tabla OPERFAB
        /// </summary>
        public DbSet<Operfab> Operfabs { get; set; }

        /// <summary>
        /// Tabla FLEXCRL
        /// </summary>
        public DbSet<FLEXCRL> Flexcrls { get; set; }

        /// <summary>
        /// Tabla MOVITRANSTO
        /// </summary>
        public DbSet<MOVITRANSTO> Movitranstos { get; set; }

        /// <summary>
        /// Tabla PEDENCA
        /// </summary>
        public DbSet<Pedenca> Pedencas { get; set; }

        /// <summary>
        /// Tabla PEDDETA
        /// </summary>
        public DbSet<Peddeta> Peddetas { get; set; }

        /// <summary>
        /// Tabla PRESUVEN
        /// </summary>
        public DbSet<Presuven> Presuvens { get; set; }

        /// <summary>
        /// Tabla CONDPAG
        /// </summary>
        public DbSet<CondPag> CondPags { get; set; }

        /// <summary>
        /// Tabla VENDEDOR
        /// </summary>
        public required DbSet<Vendedor> Vendedores { get; set; }

        /// <summary>
        /// Tabla de TRANSPORTES
        /// </summary>
        public DbSet<Transporte> Transportes { get; set; }

        /// <summary>
        /// Tabla de PedidosAuditoria
        /// </summary>
        public DbSet<PedidoAuditoria> PedidosAuditoria { get; set; }

        /// <summary>
        /// Tabla OCOMENCA - Encabezado de órdenes de compra
        /// </summary>
        public required DbSet<Ocomenca> Ocomencas { get; set; }

        /// <summary>
        /// Tabla OCOMDETA - Detalle de órdenes de compra
        /// </summary>
        public required DbSet<Ocomdeta> Ocomdetas { get; set; }

        /// <summary>
        /// Tabla DETAIVA - Detalle de IVA de órdenes de compra
        /// </summary>
        public required DbSet<Detaiva> Detaivas { get; set; }

        /// <summary>
        /// Tabla PLANCUEN - Plan de cuentas contables
        /// </summary>
        public DbSet<Plancuen> Plancuens { get; set; }

        /// <summary>
        /// Tabla MACONSIG - Consignaciones
        /// </summary>
        public DbSet<Maconsig> Maconsigs { get; set; }

        /// <summary>
        /// Tabla de Boletas de Recepción
        /// </summary>
        public DbSet<Bolrecep> Bolreceps { get; set; }

        /// <summary>
        /// Tabla de DestinoRechazoRecepcion  DESTINOS DE LOS RECHAZOS DE RECEPCION VINCULADOS CON UN DEPOSITO
        /// </summary>
        public DbSet<DestinoRechazoRecepcion> DestinoRechazoRecepciones { get; set; }

        /// <summary>
        /// Tabla de MotivosRechazoRecepcion  MOTIVOS DE RECHAZOS DE RECEPCION  
        /// </summary>
        public DbSet<MotivosRechazoRecepcion> MotivosRechazoRecepciones { get; set; }

        /// <summary>
        /// Tabla de Programas de Fabricación (Encabezados)
        /// </summary>
        public DbSet<Profabri> Profabris { get; set; }

        /// <summary>
        /// Tabla de Detalles de Programas de Fabricación
        /// </summary>
        public DbSet<Deprofab> Deprofabs { get; set; }

        /// <summary>
        /// Tabla de Cálculo de Capacidad Necesaria
        /// Almacena el detalle del cálculo de capacidad por equipo, pieza y operación
        /// </summary>
        public DbSet<Capanece> Capaneces { get; set; }
        // Resumen de cálculo de capacidad por máquina
        public DbSet<CapacidadResumenMaquina> CapacidadResumenMaquinas { get; set; }
        // Resumen de cálculo de capacidad por mano de obra
        public DbSet<CapacidadResumenManoObra> CapacidadResumenManoObras { get; set; }

        /// <summary>
        /// Tabla INDIREP
        /// </summary>
        public DbSet<Indirep> Indireps { get; set; }

        /// <summary>
        /// Tabla REPUESRE - Repuestos por Reparación
        /// </summary>
        public DbSet<Repuesre> Repuesres { get; set; }

        /// <summary>
        /// Tabla SECOPREP - Secuencia de Operaciones por Reparación
        /// </summary>
        public DbSet<Secoprep> Secopreps { get; set; }

        /// <summary>
        /// Tabla DESCREP - Descripción de Reparaciones
        /// </summary>
        public DbSet<Descrep> Descreps { get; set; }

        /// <summary>
        /// Tabla REPARA - Recepción de Bombas para Evaluar Reparación
        /// </summary>
        public DbSet<Repara> Reparas { get; set; }

        /// <summary>
        /// Tabla GALIMREP - Galería de Imágenes de Reparación
        /// </summary>
        public DbSet<GalimRep> GalimReps { get; set; }

        /// <summary>
        /// Tabla GALIMMOD - Galería de Imágenes de Modelos
        /// </summary>
        public DbSet<GalimMod> GalimMods { get; set; }

        /// <summary>
        /// Tabla DOCANRE - Documentos Anexos por Reparación (nota de reparación.)
        /// </summary>
        public DbSet<Docanre> Docanres { get; set; }

        /// <summary>
        /// Tabla REPUREPA - Repuestos por Reparación (Presupuesto/Recepción)
        /// </summary>
        public DbSet<Repurepa> Repurepas { get; set; }

        /// <summary>
        /// Tabla MASUMCLI
        /// </summary>
        public DbSet<Masumcli> Masumclis { get; set; }

        /// <summary>
        /// Tabla DETAMASU
        /// </summary>
        public DbSet<Detamasu> Detamasus { get; set; }

        /// <summary>
        /// Tabla ENTRECOM - Entregas de compras
        /// </summary>
        public DbSet<Entrecom> Entrecoms { get; set; }

        /// <summary>
        /// Tabla BADAMODE - Gestión de moldes
        /// </summary>
        public DbSet<Badamode> Badamodes { get; set; }

        /// <summary>
        /// Tabla BADAMODE_NOTA - Notas y observaciones de moldes
        /// </summary>
        public DbSet<BadamodeNota> BadamodeNotas { get; set; }

        /// <summary>
        /// Tabla DER_ACCE - Derechos de acceso de usuarios a aplicaciones
        /// </summary>
        public DbSet<DerAcce> DerAcces { get; set; }

        /// <summary>
        /// Tabla INDI_APLI - Indicadores de aplicaciones del sistema
        /// </summary>
        public DbSet<IndiApli> IndiAplis { get; set; }

        /// <summary>
        /// Tabla SyncProcesados - Auditoría de sincronización entre Central y Sucursal
        /// </summary>
        public DbSet<SyncProcesado> SyncProcesados { get; set; }

        /// <summary>
        /// Tabla SyncConfig - Configuración de sincronización entre Central y Sucursal
        /// </summary>
        public DbSet<SyncConfig> SyncConfigs { get; set; }

        #endregion

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            base.OnModelCreating(modelBuilder);
            modelBuilder.Entity<Country>().HasIndex(x => x.Name).IsUnique();
            modelBuilder.Entity<State>().HasIndex("CountryId", "Name").IsUnique();
            modelBuilder.Entity<City>().HasIndex("StateId", "Name").IsUnique();

                modelBuilder.Entity<Dependencia>().HasIndex("EmpresaId", "dependencia").IsUnique();
                //para que  no se repita si es el mismo  (crea un doble indice)
                modelBuilder.Entity<Sector>().HasIndex("DependenciaId", "sector").IsUnique();

                // Configuración de relaciones para evitar ciclos de cascada
                // Se establece DeleteBehavior.NoAction en todas las relaciones del Usuario


                // Relaciones geográficas
                modelBuilder.Entity<Usuario>()
                        .HasOne<Country>()
                        .WithMany()
                        .HasForeignKey(u => u.CountryId)
                        .OnDelete(DeleteBehavior.NoAction);

                modelBuilder.Entity<Usuario>()
                    .HasOne<State>()
                    .WithMany()
                    .HasForeignKey(u => u.StateId)
                    .OnDelete(DeleteBehavior.NoAction);

                modelBuilder.Entity<Usuario>()
                    .HasOne<City>()
                    .WithMany()
                    .HasForeignKey(u => u.CityId)
                    .OnDelete(DeleteBehavior.NoAction);

                // Relaciones organizacionales
                modelBuilder.Entity<Usuario>()
                    .HasOne<Empresa>()
                    .WithMany()
                    .HasForeignKey(u => u.EmpresaId)
                    .OnDelete(DeleteBehavior.NoAction);

                modelBuilder.Entity<Usuario>()
                    .HasOne<Dependencia>()
                    .WithMany()
                    .HasForeignKey(u => u.DependenciaId)
                    .OnDelete(DeleteBehavior.NoAction);

                modelBuilder.Entity<Usuario>()
                    .HasOne<Sector>()
                    .WithMany()
                    .HasForeignKey(u => u.SectorId)
                    .OnDelete(DeleteBehavior.NoAction);

                modelBuilder.Entity<Usuario>()
                    .HasOne<Perfil>()
                    .WithMany()
                    .HasForeignKey(u => u.PerfilId)
                    .OnDelete(DeleteBehavior.NoAction);

                modelBuilder.Entity<Tamoned>().ToTable("Tamoned");
                modelBuilder.Entity<Tamoned>()
                    .HasIndex(p => p.CodigoMoneda)
                    .IsUnique();

                modelBuilder.Entity<Familia>().ToTable("Familia");
                modelBuilder.Entity<Familia>()
                    .HasIndex(f => f.CodigoFamilia)
                    .IsUnique();

                modelBuilder.Entity<Grupiva>().ToTable("Grupiva");
                modelBuilder.Entity<Grupiva>()
                    .HasIndex(f => f.CodigoID)
                    .IsUnique();

                modelBuilder.Entity<Catoper>().ToTable("Catoper");
                modelBuilder.Entity<Catoper>()
                    .HasIndex(f => f.CodigoCategoria)
                    .IsUnique();

                modelBuilder.Entity<SecProd>().ToTable("SecProd");
                modelBuilder.Entity<SecProd>()
                    .HasIndex(f => f.CodigoSecProd)
                    .IsUnique();

                modelBuilder.Entity<PADMAQ>().ToTable("PADMAQ");
                modelBuilder.Entity<PADMAQ>()
                    .HasIndex(f => f.CodigoPadMaq)
                    .IsUnique();

                modelBuilder.Entity<Master>().ToTable("Master");
                modelBuilder.Entity<Master>()
                    .HasIndex(f => f.Codint)
                    .IsUnique();

                modelBuilder.Entity<UnidadMedidas>().ToTable("UnidadMedidas");
                modelBuilder.Entity<UnidadMedidas>()
                    .HasIndex(p => p.DescripcionAbreviada)
                    .IsUnique();

                modelBuilder.Entity<Formulas>().ToTable("Formulas");
                modelBuilder.Entity<Formulas>()
                    .HasIndex(p => p.CodiElem)
                    .IsUnique();

                modelBuilder.Entity<Secoper>().ToTable("Secoper");
                modelBuilder.Entity<Secoper>();

                // Configuración de la tabla UsuariosSectorProd
                modelBuilder.Entity<UsuariosSectorProd>().ToTable("UsuariosSectorProd");
                modelBuilder.Entity<UsuariosSectorProd>()
                    .HasIndex(u => new { u.IdFlexoft, u.CodigoSecProd })
                    .IsUnique()
                    .HasDatabaseName("IX_UsuariosSectorProd_Usuario_Sector");

                // Mapeo explícito: el DbSet Clientes apunta a la tabla física Deudor12
                modelBuilder.Entity<Deudor12>().ToTable("Deudor12");

                // Mapeo explícito: el DbSet Proveedores apunta a la tabla física Proved12
                modelBuilder.Entity<Proved12>().ToTable("Proved12");

                // Mapeo explícito: Transportes -> TRANSPORTES
                modelBuilder.Entity<Transporte>().ToTable("TRANSPORTES");
                // Índice sugerido por código de transportista si el dominio lo requiere
                modelBuilder.Entity<Transporte>()
                    .HasIndex(t => t.Codi_Tra)
                    .IsUnique(false);


                modelBuilder.Entity<PedidoAuditoria>(entity =>
                {
                    entity.ToTable("PEDIDOS_AUDITORIA");
                    entity.HasNoKey();          // <- EF solo leerá, no intentará escribir
                });

                modelBuilder.Entity<Masumcli>(entity =>
                {
                    entity.ToTable("MASUMCLI");
                    entity.HasNoKey();
                });

                // Configuración para tablas con triggers de auditoría
                // Indica a EF Core que estas tablas tienen triggers que afectan las operaciones de guardado
                // IMPORTANTE: Los nombres de triggers deben coincidir exactamente con los implementados en la BD
                modelBuilder.Entity<Pedenca>()
                    .ToTable(tb => tb.HasTrigger("TR_PEDENCA_AUDIT"));

                modelBuilder.Entity<Peddeta>()
                    .ToTable(tb => tb.HasTrigger("TR_PEDDETA_AUDIT"));

                // Configuración adicional para optimizar el rendimiento con triggers
                // Esto le dice a EF Core que recargue los datos después de INSERT/UPDATE
                // para obtener los valores modificados por los triggers
                modelBuilder.Entity<Pedenca>()
                    .Property(e => e.FechUltCambio)
                    .ValueGeneratedOnAddOrUpdate();

                modelBuilder.Entity<Peddeta>()
                    .Property(e => e.FechUltCambio)
                    .ValueGeneratedOnAddOrUpdate();
                //// Configuración de relaciones entre Pedenca y Peddeta
                //// Relación uno a muchos: Un pedido (Pedenca) puede tener múltiples detalles (Peddeta)
                //modelBuilder.Entity<Peddeta>()
                //    .HasOne<Pedenca>()
                //    .WithMany()
                //    .HasForeignKey(pd => pd.NroPed)
                //    .HasPrincipalKey(p => p.NroPed)
                //    .OnDelete(DeleteBehavior.NoAction);

                //// Relación adicional por número de cliente
                //// Un cliente puede tener múltiples pedidos y múltiples detalles de pedido
                //modelBuilder.Entity<Peddeta>()
                //    .HasOne<Pedenca>()
                //    .WithMany()
                //    .HasForeignKey(pd => pd.NroClie)
                //    .HasPrincipalKey(p => p.NroClie)
                //    .OnDelete(DeleteBehavior.NoAction);

                // Configuración de la tabla CAPANECE - Cálculo de Capacidad Necesaria
                modelBuilder.Entity<Capanece>().ToTable("CAPANECE");
                modelBuilder.Entity<Capanece>()
                    .HasIndex(c => c.Equipo)
                    .HasDatabaseName("IX_CAPANECE_Equipo");
                modelBuilder.Entity<Capanece>()
                    .HasIndex(c => c.NumeroPrograma)
                    .HasDatabaseName("IX_CAPANECE_NumeroPrograma");
                modelBuilder.Entity<Capanece>()
                    .HasIndex(c => new { c.Equipo, c.CodigoPieza })
                    .HasDatabaseName("IX_CAPANECE_Equipo_Pieza");

                modelBuilder.Entity<Indirep>()
                    .HasIndex(i => i.Codigo)
                    .IsUnique();

                modelBuilder.Entity<Repuesre>()
                    .HasIndex(r => new { r.CodigoReparacion, r.CodigoInterno })
                    .IsUnique()
                    .HasDatabaseName("IX_REPUESRE_CodigoReparacion_CodigoInterno");

                modelBuilder.Entity<Secoprep>()
                    .HasIndex(s => new { s.CodigoReparacion, s.NumeroOperacion })
                    .IsUnique()
                    .HasDatabaseName("IX_SECOPREP_CodigoReparacion_NumeroOperacion");

                // Configuración de INDI_APLI
                modelBuilder.Entity<IndiApli>()
                    .HasIndex(i => i.CodApli)
                    .IsUnique()
                    .HasDatabaseName("IX_INDI_APLI_CodApli");

                // Relación entre DER_ACCE y INDI_APLI
                // Un código de aplicación en DER_ACCE debe existir en INDI_APLI
                modelBuilder.Entity<DerAcce>()
                    .HasOne<IndiApli>()
                    .WithMany()
                    .HasForeignKey(d => d.CodApli)
                    .HasPrincipalKey(i => i.CodApli)
                    .OnDelete(DeleteBehavior.NoAction);
        }
    }
}
