# Pruebas

Ejecute desde la raiz del proyecto, con Windows PowerShell 5.1:

```powershell
powershell.exe -NoProfile -File .\tests\Probar-Cliente.ps1
powershell.exe -NoProfile -File .\tests\Probar-Receptor.ps1
```

Cliente: 11 pruebas. Receptor: 17 pruebas. Usan datos sinteticos y respuestas simuladas, sin iniciar sesion ni escribir en Excel real. Cubren la solicitud estable, el transporte, autenticacion de recepcion, cola durable, deduplicacion y confirmaciones.

En un entorno con un directorio temporal virtualizado que impida renombrar archivos, indique una carpeta escribible propia mediante `-TestRoot` a Probar-Receptor.ps1.

La suite `Probar-Inventario.ps1` conserva 17 pruebas de la alternativa anterior `Inventario-Graph.ps1`, que autentica cada computadora. Las pruebas no modifican firewall, reservas HTTP ni perfiles de Windows.
