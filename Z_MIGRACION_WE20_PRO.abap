*&---------------------------------------------------------------------*
*& Report  : Z_MIGRACION_WE20_PRO
*& Título  : Carga Masiva de Perfiles EDI (WE20)
*& Módulo  : Basis / ALE-EDI
*& Creado  : 16.03.2026
*& Autor   : FCASTELLI
*& Versión : 1.0
*&
*& Descripción:
*&   Crea de forma masiva los perfiles de interlocutor EDI leyendo un
*&   archivo Excel (.xlsx/.xls) con 20 columnas. Soporta tres tipos de
*&   registro:
*&     C – Cabecera del interlocutor (EDPP1)
*&     E – Parámetro de Entrada / Inbound (EDP21)
*&     S – Parámetro de Salida / Outbound (EDP13) + Control de Mensajes
*&         opcional (EDP12)
*&
*& Precondiciones:
*&   - El archivo Excel debe respetar exactamente el orden de las 20
*&     columnas definidas en la estructura TY_EXCEL.
*&   - La primera fila del Excel es la cabecera (se omite
*&     automáticamente con I_LINE_HEADER = 'X').
*&   - El usuario que ejecuta el reporte debe tener autorización para
*&     los objetos S_DATASET y los FM de EDI_AGREE_*.
*&
*& Columnas del Excel (en orden):
*&   1  TIPO    – Tipo de registro: C, E o S
*&   2  PARTNER – Número de interlocutor      (ej. D8430)
*&   3  PARTYP  – Tipo de interlocutor        (ej. KU)
*&   4  MATLVL  – Status de acuerdo entre interlocutores    (ej. A)
*&   5  RCVPFC  – Función del interlocutor    (ej. RE)
*&   6  MESTYP  – Tipo de mensaje EDI         (ej. INVOIC)
*&   7  MESCOD  – Código de mensaje           (ej. MM)
*&   8  MESFCT  – Función de mensaje
*&   9  TEST    – Indicador de test
*&  10  EVCODE  – Código de evento (Inbound)  (ej. INVL)
*&  11  OUTMOD  – Modo de salida              (2 ó 4)
*&  12  RCVPOR  – Puerto receptor             (ej. SAPPS4)
*&  13  IDOCTYP – Tipo de IDoc                (ej. INVOIC01)
*&  14  CIMTYP  – Tipo de extensión del IDoc
*&  15  USRTYP  – Tipo de usuario             (ej. US)
*&  16  USRKEY  – Clave de usuario / autor    (ej. CBENAVIDES)
*&  17  PCKSIZ  – Size Package    (ej. 0001)
*&  18  KAPPL   – Aplicación de MC            (ej. V3)
*&  19  KSCHL   – Clase de mensaje de MC      (ej. RD04)
*&  20  EVCODA  – Código de operación de MC   (ej. SD09)
*&---------------------------------------------------------------------*
REPORT z_migracion_we20_pro.

TYPE-POOLS: truxs.

*----------------------------------------------------------------------*
* ESTRUCTURAS DE DATOS
*----------------------------------------------------------------------*
TYPES: BEGIN OF ty_excel,
         tipo    TYPE c LENGTH 1,
         partner TYPE edpp1-parnum,
         partyp  TYPE edpp1-partyp,
         matlvl  TYPE edpp1-matlvl,
         rcvpfc  TYPE edp13-rcvpfc,
         mestyp  TYPE edp21-mestyp,
         mescod  TYPE edp21-mescod,
         mesfct  TYPE edp21-mesfct,
         test    TYPE edp21-test,
         evcode  TYPE edp21-evcode,
         outmod  TYPE edp13-outmod,
         rcvpor  TYPE edp13-rcvpor,
         idoctyp TYPE edp13-idoctyp,
         cimtyp  TYPE edp13-cimtyp,
         usrtyp  TYPE edp13-usrtyp,
         usrkey  TYPE edp13-usrkey,
         pcksiz  TYPE edp13-pcksiz,
         kappl   TYPE edp12-kappl,
         kschl   TYPE edp12-kschl,
         evcoda  TYPE edp12-evcoda,
       END OF ty_excel.

* Tabla interna y área de trabajo del Excel
DATA: gt_excel TYPE TABLE OF ty_excel,
      gs_excel TYPE ty_excel.

* Estructuras de tablas EDI
DATA: ls_edpp1 TYPE edpp1,
      ls_edp21 TYPE edp21,
      ls_edp13 TYPE edp13,
      ls_edp12 TYPE edp12.

* Buffer raw para lectura del Excel
DATA: lt_raw TYPE truxs_t_text_data.

* Contadores para el resumen final
DATA: lv_ok  TYPE i,
      lv_err TYPE i.

*----------------------------------------------------------------------*
* PANTALLA DE SELECCIÓN
*----------------------------------------------------------------------*
SELECTION-SCREEN BEGIN OF BLOCK b1 WITH FRAME TITLE TEXT-001.
  PARAMETERS: p_file TYPE rlgrap-filename OBLIGATORY.
SELECTION-SCREEN END OF BLOCK b1.

*----------------------------------------------------------------------*
* AYUDA F4 PARA EL CAMPO DE ARCHIVO
*----------------------------------------------------------------------*
AT SELECTION-SCREEN ON VALUE-REQUEST FOR p_file.
  CALL FUNCTION 'KD_GET_FILENAME_ON_F4'
    EXPORTING
      static    = 'X'
    CHANGING
      file_name = p_file.

*----------------------------------------------------------------------*
* INICIO DE PROCESAMIENTO
*----------------------------------------------------------------------*
START-OF-SELECTION.

  "--------------------------------------------------------------------
  " 1. LEER EL ARCHIVO EXCEL
  "--------------------------------------------------------------------
  CALL FUNCTION 'TEXT_CONVERT_XLS_TO_SAP'
    EXPORTING
      i_line_header        = 'X'      " Omitir fila de títulos
      i_tab_raw_data       = lt_raw
      i_filename           = p_file
    TABLES
      i_tab_converted_data = gt_excel
    EXCEPTIONS
      conversion_failed    = 1
      OTHERS               = 2.

  IF sy-subrc <> 0.
    MESSAGE 'Error leyendo el Excel. Verifique que el archivo esté cerrado e intente de nuevo.'
      TYPE 'E'.
    RETURN.
  ENDIF.

  IF gt_excel IS INITIAL.
    MESSAGE 'El archivo Excel no contiene registros de datos.' TYPE 'W'.
    RETURN.
  ENDIF.

  WRITE: / '=== INICIO DE PROCESAMIENTO WE20 ==='.
  WRITE: / 'Registros leídos:', lines( gt_excel ).
  ULINE.

  "--------------------------------------------------------------------
  " 2. PROCESAR REGISTROS
  "--------------------------------------------------------------------
  LOOP AT gt_excel INTO gs_excel.

    TRANSLATE gs_excel-tipo TO UPPER CASE.

    CASE gs_excel-tipo.

        "================================================================
        " [C] CREAR CABECERA DEL INTERLOCUTOR (EDPP1)
        "================================================================
      WHEN 'C'.
        PERFORM crear_cabecera USING gs_excel.

        "================================================================
        " [E] CREAR PARÁMETRO DE ENTRADA / INBOUND (EDP21)
        "================================================================
      WHEN 'E'.
        PERFORM crear_entrada USING gs_excel.

        "================================================================
        " [S] CREAR PARÁMETRO DE SALIDA / OUTBOUND (EDP13 + EDP12)
        "================================================================
      WHEN 'S'.
        PERFORM crear_salida USING gs_excel.

        "================================================================
        " Tipo de registro no reconocido
        "================================================================
      WHEN OTHERS.
        WRITE: / '@0A@ Fila', sy-tabix,
                 ': Tipo "', gs_excel-tipo, '" no reconocido. Se omite.'.
        ADD 1 TO lv_err.

    ENDCASE.
  ENDLOOP.

  "--------------------------------------------------------------------
  " 3. RESUMEN FINAL
  "--------------------------------------------------------------------
  ULINE.
  WRITE: / '=== RESUMEN ==='.
  WRITE: / '  Registros procesados OK   :', lv_ok.
  WRITE: / '  Registros con error/aviso :', lv_err.
  WRITE: / '=== FIN DE PROCESAMIENTO ==='.

*----------------------------------------------------------------------*
* SUBRUTINAS
*----------------------------------------------------------------------*

*&---------------------------------------------------------------------*
*& Form CREAR_CABECERA
*&   Crea la cabecera del interlocutor EDI (tabla EDPP1).
*&   Si ya existe, informa y continúa sin abortar el proceso.
*&---------------------------------------------------------------------*
FORM crear_cabecera USING ps_row TYPE ty_excel.

  CLEAR ls_edpp1.
  ls_edpp1-parnum = ps_row-partner.
  ls_edpp1-partyp = ps_row-partyp.
  ls_edpp1-matlvl = ps_row-matlvl.

  CALL FUNCTION 'EDI_AGREE_PARTNER_INSERT'
    EXPORTING
      rec_edpp1           = ls_edpp1
*     NO_PTYPE_CHECK      = ' '
    EXCEPTIONS
      db_error            = 1
      entry_already_exist = 2
      parameter_error     = 3
      OTHERS              = 4.

  IF sy-subrc = 0.
    WRITE: / '@08@ [C] CABECERA:', ps_row-partner, ps_row-partyp, '-> Creada OK.'.
    COMMIT WORK AND WAIT.
    ADD 1 TO lv_ok.
  ELSE.
    WRITE: / '@09@ [C] CABECERA:', ps_row-partner, ps_row-partyp,
             '-> Ya existía o error (sy-subrc =', sy-subrc, ').'.
    ADD 1 TO lv_err.
  ENDIF.

ENDFORM.

*&---------------------------------------------------------------------*
*& Form CREAR_ENTRADA
*&   Crea el parámetro de entrada (Inbound) para el interlocutor
*&   EDI (tabla EDP21).
*&---------------------------------------------------------------------*
FORM crear_entrada USING ps_row TYPE ty_excel.

  CLEAR ls_edp21.
  ls_edp21-sndprn = ps_row-partner.
  ls_edp21-sndprt = ps_row-partyp.
  ls_edp21-sndpfc = ps_row-rcvpfc.
  ls_edp21-mestyp = ps_row-mestyp.
  ls_edp21-mescod = ps_row-mescod.
  ls_edp21-mesfct = ps_row-mesfct.
  ls_edp21-test   = ps_row-test.
  ls_edp21-evcode = ps_row-evcode.
  ls_edp21-usrtyp = ps_row-usrtyp.
  ls_edp21-usrkey = ps_row-usrkey.


  CALL FUNCTION 'EDI_AGREE_IN_MESSTYPE_INSERT'
    EXPORTING
      rec_edp21           = ls_edp21
    EXCEPTIONS
      db_error            = 1
      entry_already_exist = 2
      parameter_error     = 3
      OTHERS              = 4.


  IF sy-subrc = 0.
    WRITE: / '@08@ [E] ENTRADA :', ps_row-partner, ps_row-mestyp, '-> Creada OK.'.
    COMMIT WORK AND WAIT.
    ADD 1 TO lv_ok.
  ELSE.
    WRITE: / '@0A@ [E] ENTRADA :', ps_row-partner, ps_row-mestyp,
             '-> Error o ya existe (sy-subrc =', sy-subrc, ').'.
    ADD 1 TO lv_err.
  ENDIF.

ENDFORM.

*&---------------------------------------------------------------------*
*& Form CREAR_SALIDA
*&   Crea el parámetro de salida (Outbound) en la tabla EDP13.
*&   Si la entrada ya existe (sy-subrc = 1) se considera válida para
*&   poder seguir con el Control de Mensajes (EDP12).
*&   Si vienen datos de Control de Mensajes (KAPPL + KSCHL), los
*&   crea en la tabla EDP12.
*&---------------------------------------------------------------------*
FORM crear_salida USING ps_row TYPE ty_excel.

  CLEAR ls_edp13.
  ls_edp13-rcvprn  = ps_row-partner.
  ls_edp13-rcvprt  = ps_row-partyp.
  ls_edp13-rcvpfc  = ps_row-rcvpfc.
  ls_edp13-mestyp  = ps_row-mestyp.
  ls_edp13-mescod  = ps_row-mescod.
  ls_edp13-mesfct  = ps_row-mesfct.
  ls_edp13-test    = ps_row-test.
  ls_edp13-outmod  = ps_row-outmod.
  ls_edp13-rcvpor  = ps_row-rcvpor.
  ls_edp13-idoctyp = ps_row-idoctyp.
  ls_edp13-cimtyp  = ps_row-cimtyp.
  ls_edp13-usrtyp  = ps_row-usrtyp.
  ls_edp13-usrkey  = ps_row-usrkey.
  ls_edp13-pcksiz  = |{ ps_row-pcksiz ALPHA = IN }|.

  CALL FUNCTION 'EDI_AGREE_OUT_MESSTYPE_INSERT'
    EXPORTING
      rec_edp13           = ls_edp13
    EXCEPTIONS
      db_error            = 1
      entry_already_exist = 2
      parameter_error     = 3
      OTHERS              = 4.

  CASE sy-subrc.
    WHEN 0.
      WRITE: / '@08@ [S] SALIDA  :', ps_row-partner, ps_row-mestyp, '-> Creada OK.'.
      COMMIT WORK AND WAIT.
      ADD 1 TO lv_ok.

    WHEN 1.
      WRITE: / '@09@ [S] SALIDA  :', ps_row-partner, ps_row-mestyp,
               '-> Ya existía (se continúa con Control de Mensajes si aplica).'.

    WHEN OTHERS.
      WRITE: / '@0A@ [S] SALIDA  :', ps_row-partner, ps_row-mestyp,
               '-> Fallo al crear (sy-subrc =', sy-subrc, '). Se omite MC.'.
      ADD 1 TO lv_err.
      RETURN.                " No crear el MC si falló la salida
  ENDCASE.

  " ---- Control de Mensajes (EDP12) – sólo si vienen KAPPL y KSCHL ----
  IF ps_row-kappl IS NOT INITIAL AND ps_row-kschl IS NOT INITIAL.

    CLEAR ls_edp12.
    ls_edp12-rcvprn = ps_row-partner.
    ls_edp12-rcvprt = ps_row-partyp.
    ls_edp12-rcvpfc = ps_row-rcvpfc.
    ls_edp12-mestyp = ps_row-mestyp.
    ls_edp12-mescod = ps_row-mescod.
    ls_edp12-mesfct = ps_row-mesfct.
    ls_edp12-test   = ps_row-test.
    ls_edp12-kappl  = ps_row-kappl.
    ls_edp12-kschl  = ps_row-kschl.
    ls_edp12-evcoda = ps_row-evcoda.

    CALL FUNCTION 'EDI_AGREE_OUT_IDOC_INSERT'
      EXPORTING
        rec_edp12           = ls_edp12
      EXCEPTIONS
        db_error            = 1
        entry_already_exist = 2
        parameter_error     = 3
        OTHERS              = 4.

    IF sy-subrc = 0.
      WRITE: / '        @0L@ -> MC (', ps_row-kappl, '/', ps_row-kschl, ') añadido OK.'.
      COMMIT WORK AND WAIT.
      ADD 1 TO lv_ok.
    ELSE.
      WRITE: / '        @0A@ -> MC (', ps_row-kappl, '/', ps_row-kschl,
               ') Error al añadir (sy-subrc =', sy-subrc, ').'.
      ADD 1 TO lv_err.
    ENDIF.

  ENDIF.

ENDFORM.