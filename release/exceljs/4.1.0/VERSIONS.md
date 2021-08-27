# OpenFin Excel API Version Numbers

## From Version 4

### Client Version

From version 4.0 onwards you can find out the client version (for log purposes etc) by logging:

```javascript
fin.desktop.Excel.version;
```

### Excel Provider Version

To confirm you are connecting to the ExcelService you expect you can log/check:

```javascript
fin.desktop.ExcelService.version;
```

### Excel Helper/Plugin Version

From version 4 you and your end users can confirm you are using the right excel plugin version by looking at Excel. 

The message will now say Connected to OpenFin (v x.x.x.x)