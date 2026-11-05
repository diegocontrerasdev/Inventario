#  Courier Control 

Sistema para una empresa de **courier que importa paquetes desde Estados Unidos**. Controla el inventario de paquetes desde que llegan a Miami hasta que se entregan al cliente, calcula automáticamente el cobro (flete, almacenaje, seguro, manejo e impuesto) y genera facturas en PDF.

> Proyecto de portafolio: demuestra cómo resolver un problema de negocio real usando solo Excel (fórmulas avanzadas) y VBA, sin depender de software externo.

---

##  Propósito

Muchos courier pequeños llevan el control de paquetes y cobros en cuadernos u hojas sueltas, lo que provoca:

- Paquetes que se pierden de vista o se entregan sin cobrar.
- Errores al calcular el peso cobrable y el almacenaje.
- Facturas hechas a mano, lentas y sin registro ordenado.

Este proyecto centraliza todo en un solo libro de Excel para **reducir errores, acelerar la facturación y tener visibilidad de lo que hay en bodega y de lo que falta por cobrar**.

Además, sirve como muestra de mis conocimientos en:

| Área | Qué demuestra este proyecto |
|---|---|
| **Excel avanzado** | `INDEX/MATCH`, `SUMIFS/COUNTIFS`, nombres definidos, validación de datos, formato condicional, áreas de impresión |
| **VBA** | Módulos organizados, constantes, validaciones, manejo de errores, creación dinámica de botones, exportación a PDF |
| **Análisis de procesos** | Modelado del flujo real de un courier (recepción → tránsito → bodega → entrega → cobro) |
| **Diseño de datos** | Hojas separadas por responsabilidad, parámetros centralizados, importes congelados al facturar |
| **Documentación** | README, guía de instalación y comentarios en el código |

---

##  Cómo funciona

### Flujo del paquete

```
Recibido en Miami → En tránsito → En bodega local → Facturado → Pagado → Entregado
```

1. **Recepción:** se registra el paquete (tracking, casillero del cliente, peso, medidas, valor declarado, ubicación).
2. **Seguimiento:** se actualiza el estado seleccionando filas; las fechas de llegada y entrega se llenan solas.
3. **Facturación:** al elegir un casillero, el sistema reúne todos sus paquetes *en bodega local* sin factura, calcula el total y genera la factura.
4. **Cobro:** se registra el pago y, opcionalmente, se marcan los paquetes como entregados.

### Reglas de cobro (configurables en la hoja `Config`)

- **Peso volumétrico** = largo × ancho × alto ÷ divisor.
- **Peso cobrable** = el mayor entre el peso real y el volumétrico, redondeado hacia arriba a libra entera.
- **Flete** = peso cobrable × tarifa por libra, con un cargo mínimo por paquete.
- **Almacenaje** = (días en bodega local − días gratis) × tarifa diaria.
- **Seguro** = porcentaje sobre el valor declarado.
- **Manejo** = cargo fijo por paquete.
- **Impuesto** = porcentaje sobre el subtotal, con total también en moneda local según el tipo de cambio.

---

##  Estructura del repositorio

```
.
├── Courier_Control.xlsx      # Libro con hojas, fórmulas y plantilla de factura
├── VBA/
│   ├── modUtil.bas           # Constantes y funciones auxiliares
│   ├── modPaquetes.bas       # Registro, búsqueda y cambio de estado
│   ├── modFacturas.bas       # Generar factura, PDF, nueva factura, registrar pago
│   ├── modSetup.bas          # Creación automática de botones
│   └── LEEME.md              # Guía rápida de instalación
└── README.md
```

### Hojas del libro

| Hoja | Función |
|---|---|
| **Panel** | Resumen: paquetes por estado, peso y valor en bodega, por cobrar y cobrado |
| **Config** | Tarifas, impuestos, tipo de cambio y datos de la empresa |
| **Clientes** | Clientes y su casillero único |
| **Recepcion** | Formulario de ingreso de paquetes con vista previa del flete |
| **Paquetes** | Inventario completo con cálculos automáticos y alertas de almacenaje |
| **Factura** | Plantilla imprimible (hasta 20 paquetes por factura) |
| **Facturas** | Registro histórico con estado de pago |

---

##  Instalación

1. Descarga o clona este repositorio.
2. Abre `Courier_Control.xlsx` en Excel (Windows o Mac).
3. Presiona **Alt + F11** para abrir el editor de VBA.
4. **Archivo → Importar archivo…** e importa los 4 archivos `.bas` de la carpeta `VBA/`.
5. Cierra el editor, presiona **Alt + F8**, ejecuta `CrearBotones`.
6. **Guardar como → Libro de Excel habilitado para macros (.xlsm)**.
7. Al reabrir el archivo, pulsa **Habilitar contenido**.

### Primer uso

1. Ajusta tarifas y datos de tu empresa en `Config`.
2. Borra los clientes y paquetes de ejemplo.
3. Registra tus clientes en `Clientes` y empieza a recibir paquetes en `Recepcion`.

---

##  Características destacadas

- Validaciones: tracking duplicado, casillero inexistente, peso obligatorio.
- Confirmación antes de emitir cada factura.
- Aviso si se intenta entregar un paquete sin facturar.
- Los importes del paquete se **congelan** al facturar para que el almacenaje no siga aumentando.
- Exportación automática de la factura a PDF en la carpeta `Facturas PDF`.
- Parámetros centralizados mediante nombres definidos: cambiar una tarifa actualiza todo el libro.

---

##  Limitaciones conocidas

- Capacidad de 500 paquetes en la tabla y 20 paquetes por factura.
- Los valores de `Config` son de ejemplo y deben ajustarse a cada negocio.
- El formato de factura no sustituye los requisitos fiscales de cada país (por ejemplo, factura electrónica); debe validarse con un contador.
- Pensado para un solo usuario a la vez (libro local).

---

##  Próximas mejoras

- [ ] Cotizador previo para clientes.
- [ ] Reporte de cuentas por cobrar por cliente.
- [ ] Envío de la factura por correo desde Excel (Outlook).
- [ ] Archivo automático de paquetes entregados.
- [ ] Versión con base de datos y acceso multiusuario.

---

##  Tecnologías

- Microsoft Excel (fórmulas, validación de datos, formato condicional)
- VBA (Visual Basic for Applications)

---



## 📄 Licencia

Este proyecto se distribuye bajo la licencia MIT. Puedes usarlo, modificarlo y adaptarlo libremente.
