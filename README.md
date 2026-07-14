## Autor

Desarrollado por Javier Rojas Castro.

## Licencia

Este proyecto está distribuido bajo la licencia MIT.

# OutlookNotifier

Complemento para **Microsoft Outlook clásico** desarrollado como una aplicación de Windows que permite generar notificaciones de audio personalizadas al recibir nuevos correos electrónicos en cualquiera de las cuentas configuradas dentro de Outlook.

El complemento monitorea las carpetas de correo en tiempo real, detectando la llegada de nuevos mensajes tanto en la **Bandeja de entrada** como en cualquier subcarpeta donde los correos sean dirigidos automáticamente mediante reglas de Outlook.

## Características principales

- 🔔 Notificación mediante audio personalizado al recibir nuevos correos.
- 📩 Compatible con múltiples cuentas configuradas en Outlook clásico.
- 📂 Monitoreo de la Bandeja de entrada y subcarpetas administradas mediante reglas de correo.
- ⚡ Detección en tiempo real utilizando los eventos nativos de Outlook.
- 🛑 Ignora carpetas del sistema como:
  - Elementos enviados.
  - Borradores.
  - Elementos eliminados.
  - Bandeja de salida.
- 🧠 Diferenciación entre correos nuevos y movimientos manuales de mensajes existentes.
- 📝 Sistema de registro (logging) para facilitar diagnóstico y soporte.
- 🔊 Reproducción de archivos de audio personalizados.

## Tecnologías utilizadas

- C#
- .NET Framework
- Visual Studio
- VSTO (Visual Studio Tools for Office)
- Microsoft Office Interop Outlook
- NAudio

## Funcionamiento

OutlookNotifier se integra como un complemento COM dentro de Outlook clásico. Al iniciar Outlook, el complemento registra las carpetas de correo disponibles y monitorea la llegada de nuevos elementos.

Cuando Outlook recibe un nuevo correo:

1. Se detecta el evento de llegada del mensaje.
2. Se valida que corresponda a un correo reciente.
3. Se verifica que no sea una carpeta excluida del sistema.
4. Se reproduce el sonido configurado como notificación.

Esto permite recibir alertas sonoras incluso cuando las reglas de Outlook distribuyen automáticamente los correos en diferentes carpetas.

## Requisitos

- Windows 10/11.
- Microsoft Outlook clásico (Microsoft 365 / Office 2024).
- .NET Framework compatible con VSTO.
- Visual Studio con herramientas de desarrollo para Office.

## Objetivo del proyecto

Crear una alternativa ligera y personalizable a las notificaciones tradicionales de Outlook, permitiendo alertas auditivas independientes del comportamiento estándar del cliente de correo y adaptadas a flujos de trabajo donde los mensajes pueden ser organizados automáticamente mediante reglas.
