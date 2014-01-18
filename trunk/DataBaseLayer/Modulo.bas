Attribute VB_Name = "Modulo"
Public TERR As New tbrErrores.clsTbrERR
Public CCFFGG As New tbrConfig.GlobalCls

Public Function GetTabla(tTablas As eTablas) As String
   GetTabla = Choose(tTablas, "Afiliado", "Alergia", "AreaProtegida", "Atencion", "Barrio", "Cargo", "Ciudad", "CodigoEmergencia", "Direccion", "Empleado", "Enfermedad", "Lugar", "Movil", "ObraSocial", "Ocupacion", "Cuota", "Parentezco", "Provincia", "ServicioEmergencia", "Sintoma" _
   , "Telefono", "TipoTelefono", "Vehiculo", "Medicamento", "TelefonoXAfiliado", "TelefonoXEmpleado", "TelefonoXAreaProtegida", "TelefonoXServicioEmergencia", "CargoXEmpleado", "TelefonoXObraSocial", "CodigoXEmpresa", "AlergiaXAfiliado", "EnfermedadXAfiliado", "MedicamentoXAfiliado" _
   , "AfiliadoExterno", "AlergiaXAfiliadoExterno", "EnfermedadXAfiliadoExterno", "MedicamentoXAfiliadoExterno", "AfiliadoExternoXObraSocial", "AfiliadoExternoXServicioEmergencia", "AfiliadoExternoXAreaProtegida", "CuotasAnuladas", "Equipo", "EmpleadoXEquipo", "EquipoXAtencion", "Recibo" _
   , "DetalleSeguimiento", "TipoCodigo", "LiquidacionEmpresa", "Guardia", "LiquidacionEmpleado", "AtencionB", "Involucrado", "InstElectrica", "InstGas", "Pais", "CuerpoBomberos", "ResponsableCuerpo", "UnidadCuerpo", "ColaboracionCuerpo")
                       
                       '    tAfiliado=1, tAlergia=2, tAreaProtegida=3, tAtencion=4,tBarrio=5 tCargo=6 tCiudad=7 tCodigoEmergencia=8 tDireccion9 tEmpleado=10 tEnfermedad11 tLugar12 tMovil13 tObraSocial14 tOcupacion15 tCuota16 tParent17  tProvincia 18  tServicioEmergencia19  Sintoma 20
   'tTelefono21 tTipoTelefono22 tVehiculo23 tMedicamento24 tTelefonoXAfiliado25 tTelefonoXEmpleado26 tTelefonoXAreaProtegida27 tTelefonoXServicioEmergencia28 tCargoXEmpleado29 tTelefonoXObraSocial30 tCodigoXObraSocial31  tAlergiaXAfiliado32 EnfermedadXAfiliado33 MedicamentoXAfiliado34
   'tAfiliadoExterno35 AlergiaXAfiliadoExterno36 EnfermedadXAfiliadoExterno37  MedicamentoXAfiliadoExterno38 tAfiliadoExternoXObraSocial39  AfiliadoExternoXServicioEmergencia40 tAfiliadoExternoXAreaProtegida41 tCuotasAnuladas42 tEquipo43   EmpleadoXEquipo44 EquipoXAtencion45  tRecibo = 46
   'seguimento=47  tipocodigo=48 liquidacionEmpresa=49 guardia=50 LiquidacionEmpleado=51, atencionB=52, involucrado=53, instElectrica=54, instGas=55, pais=56, CuerpoBomberos=57,  ResponsableCuerpo=58 , UnidadCuerpo=59 , ColaboracionCuerpo=60

End Function
