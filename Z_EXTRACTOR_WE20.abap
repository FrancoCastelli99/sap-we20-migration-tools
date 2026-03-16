*&---------------------------------------------------------------------*
*& Report  : Z_EXTRACTOR_WE20
*& Título  : Extractor Automático de Perfiles EDI (WE20) a Excel
*& Módulo  : Basis / ALE-EDI
*& Creado  : 16.03.2026
*& Autor   : FCASTELLI
*& Versión : 2.1 
*&---------------------------------------------------------------------*
REPORT z_extractor_we20.

* FIX 1: Obligatorio para poder usar SELECT-OPTIONS referenciando a la DB
TABLES: edpp1.

*----------------------------------------------------------------------*
* CLASE LOCAL - Encapsula toda la lógica del extractor
*----------------------------------------------------------------------*
CLASS lcl_extractor DEFINITION FINAL.
  PUBLIC SECTION.
    TYPES:
      " FIX 2: Definición estricta de Rangos para evitar Type Mismatches
      tt_parnum_range TYPE RANGE OF edpp1-parnum,
      tt_partyp_range TYPE RANGE OF edpp1-partyp,

      BEGIN OF ty_out,
        tipo    TYPE string,
        partner TYPE string,
        partyp  TYPE string,
        matlvl  TYPE string,
        rcvpfc  TYPE string,
        mestyp  TYPE string,
        mescod  TYPE string,
        mesfct  TYPE string,
        test    TYPE string,
        evcode  TYPE string,
        outmod  TYPE string,
        rcvpor  TYPE string,
        idoctyp TYPE string,
        cimtyp  TYPE string,
        usrtyp  TYPE string,
        usrkey  TYPE string,
        pcksiz  TYPE string,
        kappl   TYPE string,
        kschl   TYPE string,
        evcoda  TYPE string,
      END OF ty_out,
      tt_out TYPE TABLE OF ty_out WITH DEFAULT KEY.

    CLASS-METHODS:
      run
        IMPORTING iv_file   TYPE string
                  it_parnum TYPE tt_parnum_range
                  it_partyp TYPE tt_partyp_range.

  PRIVATE SECTION.
    " Tipos internos optimizados (Solo lo que se lee de BD)
    TYPES:
      BEGIN OF ty_edpp1,
        parnum TYPE edpp1-parnum,
        partyp TYPE edpp1-partyp,
        matlvl TYPE edpp1-matlvl,
      END OF ty_edpp1,

      BEGIN OF ty_edp21,
        sndprn TYPE edp21-sndprn,
        sndprt TYPE edp21-sndprt,
        sndpfc TYPE edp21-sndpfc,
        mestyp TYPE edp21-mestyp,
        mescod TYPE edp21-mescod,
        mesfct TYPE edp21-mesfct,
        test   TYPE edp21-test,
        evcode TYPE edp21-evcode,
        usrtyp TYPE edp21-usrtyp,
        usrkey TYPE edp21-usrkey,
      END OF ty_edp21,

      BEGIN OF ty_edp13,
        rcvprn TYPE edp13-rcvprn,
        rcvprt TYPE edp13-rcvprt,
        rcvpfc TYPE edp13-rcvpfc,
        mestyp TYPE edp13-mestyp,
        mescod TYPE edp13-mescod,
        mesfct TYPE edp13-mesfct,
        test   TYPE edp13-test,
        outmod TYPE edp13-outmod,
        rcvpor TYPE edp13-rcvpor,
        idoctyp TYPE edp13-idoctyp,
        cimtyp  TYPE edp13-cimtyp,
        usrtyp  TYPE edp13-usrtyp,
        usrkey  TYPE edp13-usrkey,
        pcksiz  TYPE edp13-pcksiz,
      END OF ty_edp13,

      BEGIN OF ty_edp12,
        rcvprn TYPE edp12-rcvprn,
        rcvprt TYPE edp12-rcvprt,
        rcvpfc TYPE edp12-rcvpfc,
        mestyp TYPE edp12-mestyp,
        mescod TYPE edp12-mescod,
        mesfct TYPE edp12-mesfct,
        test   TYPE edp12-test,
        kappl  TYPE edp12-kappl,
        kschl  TYPE edp12-kschl,
        evcoda TYPE edp12-evcoda,
      END OF ty_edp12,

      tt_edpp1 TYPE SORTED TABLE OF ty_edpp1  WITH UNIQUE KEY parnum partyp,
      tt_edp21 TYPE SORTED TABLE OF ty_edp21  WITH NON-UNIQUE KEY sndprn sndprt,
      tt_edp13 TYPE SORTED TABLE OF ty_edp13  WITH NON-UNIQUE KEY rcvprn rcvprt,
      tt_edp12 TYPE SORTED TABLE OF ty_edp12  WITH NON-UNIQUE KEY rcvprn rcvprt rcvpfc mestyp mescod mesfct test.

    CLASS-METHODS:
      fetch_data
        IMPORTING it_parnum        TYPE tt_parnum_range
                  it_partyp        TYPE tt_partyp_range
        EXPORTING et_edpp1         TYPE tt_edpp1
                  et_edp21         TYPE tt_edp21
                  et_edp13         TYPE tt_edp13
                  et_edp12         TYPE tt_edp12,

      build_output
        IMPORTING it_edpp1  TYPE tt_edpp1
                  it_edp21  TYPE tt_edp21
                  it_edp13  TYPE tt_edp13
                  it_edp12  TYPE tt_edp12
        RETURNING VALUE(rt_out) TYPE tt_out,

      fill_header_row
        RETURNING VALUE(rs_out) TYPE ty_out,

      fill_outbound_row
        IMPORTING is_edp13        TYPE ty_edp13
                  is_edp12        TYPE ty_edp12
        RETURNING VALUE(rs_out)   TYPE ty_out,

      download_file
        IMPORTING iv_file  TYPE string
                  it_out   TYPE tt_out.

ENDCLASS.

CLASS lcl_extractor IMPLEMENTATION.

  METHOD run.
    DATA(lt_edpp1) = VALUE tt_edpp1( ).
    DATA(lt_edp21) = VALUE tt_edp21( ).
    DATA(lt_edp13) = VALUE tt_edp13( ).
    DATA(lt_edp12) = VALUE tt_edp12( ).

    fetch_data(
      EXPORTING
        it_parnum = it_parnum
        it_partyp = it_partyp
      IMPORTING
        et_edpp1  = lt_edpp1
        et_edp21  = lt_edp21
        et_edp13  = lt_edp13
        et_edp12  = lt_edp12
    ).

    IF lt_edpp1 IS INITIAL.
      " FIX 3: Mensaje directo y claro en lugar de TEXT-e01
      MESSAGE 'No se encontraron interlocutores con esos filtros.' TYPE 'I'.
      RETURN.
    ENDIF.

    DATA(lt_out) = build_output(
      it_edpp1 = lt_edpp1
      it_edp21 = lt_edp21
      it_edp13 = lt_edp13
      it_edp12 = lt_edp12
    ).

    download_file( iv_file = iv_file  it_out = lt_out ).
  ENDMETHOD.

  METHOD fetch_data.
    SELECT parnum, partyp, matlvl
      FROM edpp1
      INTO TABLE @et_edpp1
      WHERE parnum IN @it_parnum
        AND partyp IN @it_partyp.

    CHECK et_edpp1 IS NOT INITIAL.

    SELECT sndprn, sndprt, sndpfc, mestyp, mescod, mesfct, test, evcode, usrtyp, usrkey
      FROM edp21
      INTO TABLE @et_edp21
      FOR ALL ENTRIES IN @et_edpp1
      WHERE sndprn = @et_edpp1-parnum
        AND sndprt = @et_edpp1-partyp.

    SELECT rcvprn, rcvprt, rcvpfc, mestyp, mescod, mesfct, test,
           outmod, rcvpor, idoctyp, cimtyp, usrtyp, usrkey, pcksiz
      FROM edp13
      INTO TABLE @et_edp13
      FOR ALL ENTRIES IN @et_edpp1
      WHERE rcvprn = @et_edpp1-parnum
        AND rcvprt = @et_edpp1-partyp.

    CHECK et_edp13 IS NOT INITIAL.

    SELECT rcvprn, rcvprt, rcvpfc, mestyp, mescod, mesfct, test, kappl, kschl, evcoda
      FROM edp12
      INTO TABLE @et_edp12
      FOR ALL ENTRIES IN @et_edp13
      WHERE rcvprn = @et_edp13-rcvprn
        AND rcvprt = @et_edp13-rcvprt
        AND rcvpfc = @et_edp13-rcvpfc
        AND mestyp = @et_edp13-mestyp
        AND mescod = @et_edp13-mescod
        AND mesfct = @et_edp13-mesfct
        AND test   = @et_edp13-test.
  ENDMETHOD.

  METHOD build_output.
    APPEND fill_header_row( ) TO rt_out.

    LOOP AT it_edpp1 INTO DATA(ls_edpp1).

      "[C] Cabecera
      APPEND VALUE ty_out(
        tipo    = 'C'
        partner = ls_edpp1-parnum
        partyp  = ls_edpp1-partyp
        matlvl  = ls_edpp1-matlvl
      ) TO rt_out.

      "[E] Entradas
      LOOP AT it_edp21 INTO DATA(ls_edp21) WHERE sndprn = ls_edpp1-parnum AND sndprt = ls_edpp1-partyp.
        APPEND VALUE ty_out(
          tipo    = 'E'
          partner = ls_edp21-sndprn
          partyp  = ls_edp21-sndprt
          rcvpfc  = ls_edp21-sndpfc
          mestyp  = ls_edp21-mestyp
          mescod  = ls_edp21-mescod
          mesfct  = ls_edp21-mesfct
          test    = ls_edp21-test
          evcode  = ls_edp21-evcode
          usrtyp  = ls_edp21-usrtyp
          usrkey  = ls_edp21-usrkey
        ) TO rt_out.
      ENDLOOP.

      " [S] Salidas + Message Control
      LOOP AT it_edp13 INTO DATA(ls_edp13) WHERE rcvprn = ls_edpp1-parnum AND rcvprt = ls_edpp1-partyp.

        DATA(lv_has_mc) = abap_false.

        LOOP AT it_edp12 INTO DATA(ls_edp12)
          WHERE rcvprn = ls_edp13-rcvprn
            AND rcvprt = ls_edp13-rcvprt
            AND rcvpfc = ls_edp13-rcvpfc
            AND mestyp = ls_edp13-mestyp
            AND mescod = ls_edp13-mescod
            AND mesfct = ls_edp13-mesfct
            AND test   = ls_edp13-test.

          lv_has_mc = abap_true.
          APPEND fill_outbound_row( is_edp13 = ls_edp13  is_edp12 = ls_edp12 ) TO rt_out.
        ENDLOOP.

        " Sin Message Control
        IF lv_has_mc = abap_false.
          APPEND fill_outbound_row( is_edp13 = ls_edp13  is_edp12 = VALUE #( ) ) TO rt_out.
        ENDIF.

      ENDLOOP.
    ENDLOOP.
  ENDMETHOD.

  METHOD fill_header_row.
    rs_out = VALUE ty_out(
      tipo    = 'TIPO'    partner = 'PARTNER' partyp  = 'PARTYP'
      matlvl  = 'MATLVL'  rcvpfc  = 'RCVPFC'  mestyp  = 'MESTYP'
      mescod  = 'MESCOD'  mesfct  = 'MESFCT'  test    = 'TEST'
      evcode  = 'EVCODE'  outmod  = 'OUTMOD'  rcvpor  = 'RCVPOR'
      idoctyp = 'IDOCTYP' cimtyp  = 'CIMTYP'  usrtyp  = 'USRTYP'
      usrkey  = 'USRKEY'  pcksiz  = 'PCKSIZ'  kappl   = 'KAPPL'
      kschl   = 'KSCHL'   evcoda  = 'EVCODA'
    ).
  ENDMETHOD.

  METHOD fill_outbound_row.
    rs_out = VALUE ty_out(
      tipo    = 'S'
      partner = is_edp13-rcvprn
      partyp  = is_edp13-rcvprt
      rcvpfc  = is_edp13-rcvpfc
      mestyp  = is_edp13-mestyp
      mescod  = is_edp13-mescod
      mesfct  = is_edp13-mesfct
      test    = is_edp13-test
      outmod  = is_edp13-outmod
      rcvpor  = is_edp13-rcvpor
      idoctyp = is_edp13-idoctyp
      cimtyp  = is_edp13-cimtyp
      usrtyp  = is_edp13-usrtyp
      usrkey  = is_edp13-usrkey
    ).

    rs_out-pcksiz = is_edp13-pcksiz.
    CONDENSE rs_out-pcksiz.
    SHIFT rs_out-pcksiz LEFT DELETING LEADING '0'.

    IF is_edp12-kappl IS NOT INITIAL.
      rs_out-kappl  = is_edp12-kappl.
      rs_out-kschl  = is_edp12-kschl.
      rs_out-evcoda = is_edp12-evcoda.
    ENDIF.
  ENDMETHOD.

  METHOD download_file.
    CALL FUNCTION 'GUI_DOWNLOAD'
      EXPORTING
        filename              = iv_file
        filetype              = 'ASC'
        write_field_separator = abap_true
      TABLES
        data_tab              = it_out
      EXCEPTIONS
        OTHERS                = 1.

    IF sy-subrc = 0.
      MESSAGE |Extracción OK: { lines( it_out ) } registros generados en el Excel.| TYPE 'I'.
    ELSE.
      MESSAGE 'Error al descargar el archivo en su PC.' TYPE 'E'.
    ENDIF.
  ENDMETHOD.

ENDCLASS.

*----------------------------------------------------------------------*
* PANTALLA DE SELECCIÓN
*----------------------------------------------------------------------*
SELECTION-SCREEN BEGIN OF BLOCK b1 WITH FRAME TITLE TEXT-001.
  SELECT-OPTIONS: s_parnum FOR edpp1-parnum,
                  s_partyp FOR edpp1-partyp.
SELECTION-SCREEN END OF BLOCK b1.

SELECTION-SCREEN BEGIN OF BLOCK b2 WITH FRAME TITLE TEXT-002.
  PARAMETERS: p_file TYPE string LOWER CASE OBLIGATORY
              DEFAULT 'C:\Temp\WE20_Export.xls'.
SELECTION-SCREEN END OF BLOCK b2.

AT SELECTION-SCREEN ON VALUE-REQUEST FOR p_file.
  DATA: lv_filename TYPE string,
        lv_path     TYPE string,
        lv_fullpath TYPE string.

  cl_gui_frontend_services=>file_save_dialog(
    EXPORTING
      window_title      = 'Guardar Excel Extractor WE20'
      default_extension = 'xls'
      default_file_name = 'WE20_Export.xls'
    CHANGING
      filename          = lv_filename
      path              = lv_path
      fullpath          = lv_fullpath
    EXCEPTIONS
      OTHERS            = 1
  ).
  IF sy-subrc = 0.
    p_file = lv_fullpath.
  ENDIF.

*----------------------------------------------------------------------*
* PUNTO DE ENTRADA
*----------------------------------------------------------------------*
START-OF-SELECTION.
  lcl_extractor=>run(
    iv_file   = p_file
    it_parnum = s_parnum[]
    it_partyp = s_partyp[]
    ).