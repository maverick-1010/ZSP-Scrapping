*&---------------------------------------------------------------------*
*& Report  ZMMSPARESCRAP
*& Tcode   ZSCRAP
*& Title   Spares Scrap Upload
*&---------------------------------------------------------------------*
*& Demand  : DMND0025000 - ZSPF Scrapping process for spares plants
*& Purpose : Scrap inventory of the Spares plants through a controlled
*&           Excel upload.
*&             Unrestricted : 344 (Unres -> Blocked) then 555 (scrap)
*&             QI stock     : 350 (QI    -> Blocked) then 555 (scrap)
*&             Blocked      : 555 (scrap)
*&           If 555 fails after a 344/350 posted, the transfer is
*&           reversed with 343 / 349 (agreed approach).
*&
*& Single-file delivery. Sections are separated by banners that mirror
*& the includes the program can be split into later:
*&   ZMMSPARESCRAP_TOP  - types, constants, global data
*&   ZMMSPARESCRAP_SEL  - selection screen
*&   ZMMSPARESCRAP_CLS  - local class definition + implementation
*&   ZMMSPARESCRAP_MAIN - events
*&
*& Prerequisites (see README.md):
*&   - Number range object ZSCRAP, interval 01 (0000000-9999999)
*&   - TVARVC  ZSCRAP_PLANT (type S, one row per allowed plant)
*&   - Tcode ZSCRAP -> this report; authorization to be assigned
*&   - BASIS: SAPoffice/GOS storage for attachments
*&---------------------------------------------------------------------*
REPORT zmmsparescrap.

*&---------------------------------------------------------------------*
*& INCLUDE ZMMSPARESCRAP_TOP
*&---------------------------------------------------------------------*
TYPES: BEGIN OF ty_up,
         row    TYPE i,
         matnr  TYPE matnr,
         maktx  TYPE maktx,
         werks  TYPE werks_d,
         lgort  TYPE lgort_d,
         meins  TYPE meins,
         unres  TYPE menge_d,
         qi     TYPE menge_d,
         blk    TYPE menge_d,
         disp   TYPE menge_d,
         s_disp TYPE string,          " raw text of Disposal Qty (blank check)
         value  TYPE p LENGTH 13 DECIMALS 2,
         reason TYPE c LENGTH 50,
         charg  TYPE charg_d,
         xchpf  TYPE xchpf,           " batch managed (MARA)
         bwtar  TYPE bwtar_d,         " valuation type (MCHA)
         err    TYPE abap_bool,
       END OF ty_up,
       tt_up TYPE STANDARD TABLE OF ty_up WITH DEFAULT KEY.

TYPES: BEGIN OF ty_log,
         status  TYPE c LENGTH 60,
         mblnr   TYPE mblnr,
         matnr   TYPE matnr,
         bwart   TYPE bwart,
         menge   TYPE menge_d,
         werks   TYPE werks_d,
         lgort   TYPE lgort_d,
         charg   TYPE charg_d,
         message TYPE c LENGTH 220,
         row     TYPE i,
         mjahr   TYPE mjahr,
         rowcol  TYPE c LENGTH 4,
       END OF ty_log,
       tt_log TYPE STANDARD TABLE OF ty_log WITH DEFAULT KEY.

CONSTANTS: gc_tvarvc_name TYPE tvarvc-name VALUE 'ZSCRAP_PLANT',
           gc_nr_object   TYPE nrobj       VALUE 'ZSCRAP',
           gc_nr_range    TYPE nrnr        VALUE '01',
           gc_costcenter  TYPE kostl       VALUE '0000003127',
           gc_reason      TYPE mb_grbew    VALUE '0004',
           gc_sheet       TYPE string      VALUE 'Upload File',
           gc_ok          TYPE c LENGTH 60 VALUE 'Success',
           gc_err         TYPE c LENGTH 60 VALUE 'Error - Review and reupload this record only.'.

*&---------------------------------------------------------------------*
*& INCLUDE ZMMSPARESCRAP_SEL
*&---------------------------------------------------------------------*
DATA gv_smtp TYPE ad_smtpadr.

SELECTION-SCREEN BEGIN OF BLOCK b1 WITH FRAME TITLE t_b1.
PARAMETERS:     p_file TYPE c LENGTH 255 LOWER CASE OBLIGATORY,
                p_mail TYPE c LENGTH 255 LOWER CASE OBLIGATORY.
SELECT-OPTIONS: s_email FOR gv_smtp NO INTERVALS.
PARAMETERS:     p_test AS CHECKBOX DEFAULT 'X'.
SELECTION-SCREEN END OF BLOCK b1.

*&---------------------------------------------------------------------*
*& INCLUDE ZMMSPARESCRAP_CLS  (definition)
*&---------------------------------------------------------------------*
CLASS lcl_scrap DEFINITION FINAL.
  PUBLIC SECTION.
    CLASS-METHODS:
      f4_file IMPORTING iv_filter TYPE string
              CHANGING  cv_file   TYPE c.

    METHODS:
      constructor IMPORTING iv_file TYPE c
                            iv_mail TYPE c
                            iv_test TYPE abap_bool,
      run,
      read_files,
      validate    RETURNING VALUE(rv_ok) TYPE abap_bool,
      process,
      attach_email,
      show_alv,
      send_mail.

  PRIVATE SECTION.
    DATA: mv_file    TYPE string,
          mv_mail    TYPE string,
          mv_test    TYPE abap_bool,
          mv_xlsx    TYPE xstring,       " upload file content
          mt_mail    TYPE solix_tab,     " approval e-mail content
          mv_maillen TYPE i,
          mv_number  TYPE n LENGTH 7,    " ZSCRAP number of this upload
          mt_up      TYPE tt_up,
          mt_log     TYPE tt_log.

    METHODS:
      read_binary  IMPORTING iv_path    TYPE string
                   EXPORTING et_solix   TYPE solix_tab
                             ev_length  TYPE i
                             ev_ok      TYPE abap_bool,
      parse_excel,
      conv_num     IMPORTING iv_txt TYPE string
                   RETURNING VALUE(rv_num) TYPE f
                   RAISING   cx_sy_conversion_error,
      add_log      IMPORTING iv_ok   TYPE abap_bool
                             iv_msg  TYPE string
                             iv_row  TYPE i OPTIONAL
                             iv_matnr TYPE matnr OPTIONAL
                             iv_bwart TYPE bwart OPTIONAL
                             iv_menge TYPE menge_d OPTIONAL
                             iv_werks TYPE werks_d OPTIONAL
                             iv_lgort TYPE lgort_d OPTIONAL
                             iv_charg TYPE charg_d OPTIONAL
                             iv_mblnr TYPE mblnr OPTIONAL
                             iv_mjahr TYPE mjahr OPTIONAL,
      mat_out      IMPORTING iv_matnr TYPE matnr
                   RETURNING VALUE(rv_txt) TYPE string,
      get_plant_range RETURNING VALUE(rt_range) TYPE rseloption,
      process_row  IMPORTING is_row TYPE ty_up,
      post_seq     IMPORTING is_row  TYPE ty_up
                             iv_xfer TYPE bwart
                             iv_rev  TYPE bwart
                             iv_qty  TYPE menge_d
                   RETURNING VALUE(rv_ok) TYPE abap_bool,
      post_gm      IMPORTING is_row  TYPE ty_up
                             iv_bwart TYPE bwart
                             iv_qty  TYPE menge_d
                             iv_note TYPE string OPTIONAL
                   RETURNING VALUE(rv_ok) TYPE abap_bool,
      build_salv   IMPORTING iv_list TYPE abap_bool
                   EXPORTING eo_salv TYPE REF TO cl_salv_table
                   CHANGING  ct_log TYPE tt_log.
ENDCLASS.

*&---------------------------------------------------------------------*
*& INCLUDE ZMMSPARESCRAP_CLS  (implementation)
*&---------------------------------------------------------------------*
CLASS lcl_scrap IMPLEMENTATION.

  METHOD f4_file.
    DATA: lt_files TYPE filetable,
          lv_rc    TYPE i.

    cl_gui_frontend_services=>file_open_dialog(
      EXPORTING file_filter = iv_filter
      CHANGING  file_table  = lt_files
                rc          = lv_rc
      EXCEPTIONS OTHERS     = 1 ).
    IF sy-subrc = 0 AND lv_rc = 1.
      cv_file = lt_files[ 1 ]-filename.
    ENDIF.
  ENDMETHOD.

  METHOD constructor.
    mv_file = iv_file.
    mv_mail = iv_mail.
    mv_test = iv_test.
  ENDMETHOD.

  "--------------------------------------------------------------------
  " Main flow: read -> validate -> (test run stop) -> process -> output
  "--------------------------------------------------------------------
  METHOD run.
    read_files( ).

    IF mt_log IS INITIAL AND validate( ) = abap_true.
      IF mv_test = abap_true.
        MESSAGE 'File can be uploaded. Please untick Test Run' TYPE 'S'.
        RETURN.
      ENDIF.
      process( ).
      attach_email( ).
    ENDIF.

    show_alv( ).
    IF mv_test = abap_false.
      send_mail( ).
    ENDIF.
  ENDMETHOD.

  "--------------------------------------------------------------------
  " Read Excel + approval e-mail from the frontend (GUI and WebGUI)
  "--------------------------------------------------------------------
  METHOD read_binary.
    DATA lt_raw TYPE solix_tab.

    ev_ok = abap_false.
    cl_gui_frontend_services=>gui_upload(
      EXPORTING filename   = iv_path
                filetype   = 'BIN'
      IMPORTING filelength = ev_length
      CHANGING  data_tab   = lt_raw
      EXCEPTIONS OTHERS    = 1 ).
    IF sy-subrc <> 0.
      RETURN.
    ENDIF.
    et_solix = lt_raw.
    ev_ok = abap_true.
  ENDMETHOD.

  METHOD read_files.
    DATA: lt_x   TYPE solix_tab,
          lv_len TYPE i,
          lv_ok  TYPE abap_bool.

    " Approval e-mail: check extension first (scenario 15)
    DATA(lv_low) = to_lower( mv_mail ).
    IF NOT ( lv_low CP '*.msg' OR lv_low CP '*.pdf' ).
      add_log( iv_ok = abap_false
               iv_msg = 'Please upload MSG or PDF format approval email' ).
    ENDIF.

    " Excel
    read_binary( EXPORTING iv_path = mv_file
                 IMPORTING et_solix = lt_x
                           ev_length = lv_len
                           ev_ok = lv_ok ).
    IF lv_ok = abap_false.
      add_log( iv_ok = abap_false
               iv_msg = |Upload file { mv_file } could not be read| ).
      RETURN.
    ENDIF.
    CALL FUNCTION 'SCMS_BINARY_TO_XSTRING'
      EXPORTING
        input_length = lv_len
      IMPORTING
        buffer       = mv_xlsx
      TABLES
        binary_tab   = lt_x
      EXCEPTIONS
        failed       = 1
        OTHERS       = 2.
    IF sy-subrc <> 0.
      add_log( iv_ok = abap_false iv_msg = 'Upload file could not be converted' ).
      RETURN.
    ENDIF.

    " Approval e-mail content (only needed for a live run)
    IF mv_test = abap_false AND mt_log IS INITIAL.
      read_binary( EXPORTING iv_path = mv_mail
                   IMPORTING et_solix = mt_mail
                             ev_length = mv_maillen
                             ev_ok = lv_ok ).
      IF lv_ok = abap_false.
        add_log( iv_ok = abap_false
                 iv_msg = |Approval email { mv_mail } could not be read| ).
      ENDIF.
    ENDIF.

    IF mv_xlsx IS NOT INITIAL.
      parse_excel( ).
    ENDIF.
  ENDMETHOD.

  METHOD conv_num.
    DATA(lv_txt) = condense( iv_txt ).
    IF lv_txt IS INITIAL.
      rv_num = 0.
      RETURN.
    ENDIF.
    rv_num = lv_txt.      " raises CX_SY_CONVERSION_NO_NUMBER if not numeric
  ENDMETHOD.

  METHOD parse_excel.
    DATA: lo_data TYPE REF TO data,
          lv_row  TYPE i,
          lv_off  TYPE i VALUE 1,
          ls_up   TYPE ty_up.
    FIELD-SYMBOLS: <lt_data> TYPE STANDARD TABLE,
                   <ls_row>  TYPE any,
                   <lv_col>  TYPE any.

    TRY.
        DATA(lo_xl) = NEW cl_fdt_xl_spreadsheet( document_name = 'upload.xlsx'
                                                 xdocument     = mv_xlsx ).
        lo_xl->if_fdt_doc_spreadsheet~get_worksheet_names(
          IMPORTING worksheet_names = DATA(lt_sheets) ).
        IF lt_sheets IS INITIAL.
          add_log( iv_ok = abap_false iv_msg = 'Upload file has no worksheet' ).
          RETURN.
        ENDIF.
        " Sheet 'Upload File' if present, else first sheet
        DATA(lv_sheet) = COND string( WHEN line_exists( lt_sheets[ table_line = gc_sheet ] )
                                      THEN gc_sheet
                                      ELSE lt_sheets[ 1 ] ).
        lo_data = lo_xl->if_fdt_doc_spreadsheet~get_itab_from_worksheet( lv_sheet ).
      CATCH cx_root INTO DATA(lx).
        add_log( iv_ok = abap_false
                 iv_msg = |Upload file is not a valid Excel (xlsx): { lx->get_text( ) }| ).
        RETURN.
    ENDTRY.

    ASSIGN lo_data->* TO <lt_data>.
    IF <lt_data> IS NOT ASSIGNED.
      RETURN.
    ENDIF.

    " Columns by position: 1 Material, 2 Description, 3 Plant, 4 Sloc,
    " 5 UoM, 6 UNREST, 7 QI, 8 BLK, 9 Disposal Qty, 10 $Value,
    " 11 Reason, 12 Batch (anything further, e.g. col M notes, ignored)
    LOOP AT <lt_data> ASSIGNING <ls_row>.
      lv_row = sy-tabix.
      CLEAR ls_up.
      ls_up-row = lv_row + lv_off.     " Excel row (header = row 1)

      DATA(lv_txt) = ||.
      DATA(lv_skip) = abap_false.
      DO 12 TIMES.
        DATA(lv_idx) = sy-index.
        ASSIGN COMPONENT lv_idx OF STRUCTURE <ls_row> TO <lv_col>.
        IF sy-subrc <> 0.
          EXIT.
        ENDIF.
        lv_txt = condense( CONV string( <lv_col> ) ).
        CASE lv_idx.
          WHEN 1.
            IF to_upper( lv_txt ) = 'MATERIAL'.   " header row, if returned
              lv_skip = abap_true.
              lv_off  = 0.
              EXIT.
            ENDIF.
            ls_up-matnr = to_upper( lv_txt ).
            IF ls_up-matnr IS NOT INITIAL.
              CALL FUNCTION 'CONVERSION_EXIT_MATN1_INPUT'
                EXPORTING  input        = ls_up-matnr
                IMPORTING  output       = ls_up-matnr
                EXCEPTIONS OTHERS       = 1.
            ENDIF.
          WHEN 2.  ls_up-maktx = lv_txt.
          WHEN 3.  ls_up-werks = to_upper( lv_txt ).
          WHEN 4.  ls_up-lgort = to_upper( lv_txt ).
          WHEN 5.  ls_up-meins = to_upper( lv_txt ).
          WHEN 6 OR 7 OR 8 OR 9 OR 10.
            TRY.
                DATA(lv_num) = conv_num( lv_txt ).
                CASE lv_idx.
                  WHEN 6.  ls_up-unres  = lv_num.
                  WHEN 7.  ls_up-qi     = lv_num.
                  WHEN 8.  ls_up-blk    = lv_num.
                  WHEN 9.  ls_up-disp   = lv_num.
                            ls_up-s_disp = lv_txt.
                  WHEN 10. ls_up-value  = lv_num.
                ENDCASE.
              CATCH cx_sy_conversion_error.
                ls_up-err = abap_true.
                add_log( iv_ok = abap_false iv_row = ls_up-row
                         iv_matnr = ls_up-matnr
                         iv_msg = |Row { ls_up-row }: column { lv_idx } '{ lv_txt }' is not a number| ).
            ENDTRY.
          WHEN 11. ls_up-reason = lv_txt.
          WHEN 12. ls_up-charg  = to_upper( lv_txt ).
        ENDCASE.
      ENDDO.

      IF lv_skip = abap_true.
        CONTINUE.
      ENDIF.
      " completely blank line
      IF ls_up-matnr IS INITIAL AND ls_up-werks IS INITIAL AND ls_up-lgort IS INITIAL.
        CONTINUE.
      ENDIF.
      APPEND ls_up TO mt_up.
    ENDLOOP.

    IF mt_up IS INITIAL AND mt_log IS INITIAL.
      add_log( iv_ok = abap_false iv_msg = 'Upload file contains no records' ).
    ENDIF.
  ENDMETHOD.

  "--------------------------------------------------------------------
  " Helpers
  "--------------------------------------------------------------------
  METHOD mat_out.
    rv_txt = |{ iv_matnr ALPHA = OUT }|.
    rv_txt = condense( rv_txt ).
  ENDMETHOD.

  METHOD add_log.
    APPEND VALUE ty_log(
      status  = COND #( WHEN iv_ok = abap_true THEN gc_ok ELSE gc_err )
      mblnr   = iv_mblnr
      mjahr   = iv_mjahr
      matnr   = iv_matnr
      bwart   = iv_bwart
      menge   = iv_menge
      werks   = iv_werks
      lgort   = iv_lgort
      charg   = iv_charg
      message = iv_msg
      row     = iv_row
      rowcol  = COND #( WHEN iv_ok = abap_true THEN 'C510' ELSE 'C610' ) ) TO mt_log.
  ENDMETHOD.

  METHOD get_plant_range.
    SELECT sign, opti, low, high FROM tvarvc
      WHERE name = @gc_tvarvc_name
        AND type IN ('S', 'P')
      INTO TABLE @DATA(lt_tv).
    LOOP AT lt_tv ASSIGNING FIELD-SYMBOL(<ls_tv>).
      APPEND VALUE #( sign   = COND #( WHEN <ls_tv>-sign IS INITIAL THEN 'I' ELSE <ls_tv>-sign )
                      option = COND #( WHEN <ls_tv>-opti IS INITIAL THEN 'EQ' ELSE <ls_tv>-opti )
                      low    = <ls_tv>-low
                      high   = <ls_tv>-high ) TO rt_range.
    ENDLOOP.
  ENDMETHOD.

  "--------------------------------------------------------------------
  " Validation - whole file, nothing is posted if any error is found
  "--------------------------------------------------------------------
  METHOD validate.
    DATA lv_sum TYPE menge_d.

    IF mt_up IS INITIAL.
      RETURN.
    ENDIF.

    " ---- bulk selects ------------------------------------------------
    SELECT matnr, meins, xchpf FROM mara
      FOR ALL ENTRIES IN @mt_up
      WHERE matnr = @mt_up-matnr
      INTO TABLE @DATA(lt_mara).

    SELECT matnr, werks, lgort FROM mard
      FOR ALL ENTRIES IN @mt_up
      WHERE matnr = @mt_up-matnr
        AND werks = @mt_up-werks
        AND lgort = @mt_up-lgort
      INTO TABLE @DATA(lt_mard).

    SELECT matnr, werks, lgort, charg, clab, cinsm, cspem FROM mchb
      FOR ALL ENTRIES IN @mt_up
      WHERE matnr = @mt_up-matnr
        AND werks = @mt_up-werks
        AND lgort = @mt_up-lgort
        AND charg = @mt_up-charg
      INTO TABLE @DATA(lt_mchb).

    SELECT matnr, werks, charg, bwtar FROM mcha
      FOR ALL ENTRIES IN @mt_up
      WHERE matnr = @mt_up-matnr
        AND werks = @mt_up-werks
        AND charg = @mt_up-charg
      INTO TABLE @DATA(lt_mcha).

    DATA(lt_plants) = get_plant_range( ).
    DATA lt_plant_err TYPE SORTED TABLE OF werks_d WITH UNIQUE KEY table_line.

    " ---- row checks ----------------------------------------------------
    LOOP AT mt_up ASSIGNING FIELD-SYMBOL(<up>).
      DATA(lv_mat) = mat_out( <up>-matnr ).

      " 1. mandatory columns
      DATA(lt_miss) = VALUE string_table( ).
      IF <up>-matnr  IS INITIAL. APPEND 'Material'        TO lt_miss. ENDIF.
      IF <up>-werks  IS INITIAL. APPEND 'Plant'           TO lt_miss. ENDIF.
      IF <up>-lgort  IS INITIAL. APPEND 'Sloc'            TO lt_miss. ENDIF.
      IF <up>-meins  IS INITIAL. APPEND 'UOM'             TO lt_miss. ENDIF.
      IF <up>-s_disp IS INITIAL. APPEND 'Disposal Qty'    TO lt_miss. ENDIF.
      IF <up>-reason IS INITIAL. APPEND 'Reason For Disposal' TO lt_miss. ENDIF.
      IF <up>-charg  IS INITIAL. APPEND 'Batch Number'    TO lt_miss. ENDIF.
      LOOP AT lt_miss INTO DATA(lv_col).
        add_log( iv_ok = abap_false iv_row = <up>-row iv_matnr = <up>-matnr
                 iv_werks = <up>-werks iv_lgort = <up>-lgort
                 iv_msg = |Row { <up>-row }: Please ensure { lv_col } is filled up| ).
        <up>-err = abap_true.
      ENDLOOP.
      IF <up>-matnr IS INITIAL.
        CONTINUE.
      ENDIF.

      " 2. material exists (MARA)
      READ TABLE lt_mara ASSIGNING FIELD-SYMBOL(<mara>)
           WITH KEY matnr = <up>-matnr.
      IF sy-subrc <> 0.
        add_log( iv_ok = abap_false iv_row = <up>-row iv_matnr = <up>-matnr
                 iv_msg = |Material { lv_mat } does not exists.| ).
        <up>-err = abap_true.
        CONTINUE.
      ENDIF.
      <up>-xchpf = <mara>-xchpf.

      " 3. plant allowed (TVARVC ZSCRAP_PLANT)
      IF <up>-werks IS NOT INITIAL AND ( lt_plants IS INITIAL OR <up>-werks NOT IN lt_plants ).
        IF NOT line_exists( lt_plant_err[ table_line = <up>-werks ] ).
          INSERT <up>-werks INTO TABLE lt_plant_err.
          add_log( iv_ok = abap_false iv_row = <up>-row iv_werks = <up>-werks
                   iv_msg = |Plant { <up>-werks } is not allowed. Please maintain TVARVC| ).
        ENDIF.
        <up>-err = abap_true.
      ENDIF.

      " 4. extended to plant / sloc (MARD)
      IF <up>-werks IS NOT INITIAL AND <up>-lgort IS NOT INITIAL
         AND NOT line_exists( lt_mard[ matnr = <up>-matnr werks = <up>-werks
                                       lgort = <up>-lgort ] ).
        add_log( iv_ok = abap_false iv_row = <up>-row iv_matnr = <up>-matnr
                 iv_werks = <up>-werks iv_lgort = <up>-lgort
                 iv_msg = |Material { lv_mat } not extended to { <up>-werks } { <up>-lgort }| ).
        <up>-err = abap_true.
      ENDIF.

      " 5. quantities: Disposal = Unres + QI + Blocked
      lv_sum = <up>-unres + <up>-qi + <up>-blk.
      IF <up>-s_disp IS NOT INITIAL AND lv_sum <> <up>-disp.
        add_log( iv_ok = abap_false iv_row = <up>-row iv_matnr = <up>-matnr
                 iv_msg = |Mismatch quantities in material { lv_mat }| ).
        <up>-err = abap_true.
      ENDIF.
      IF lv_sum <= 0 AND <up>-s_disp IS NOT INITIAL.
        add_log( iv_ok = abap_false iv_row = <up>-row iv_matnr = <up>-matnr
                 iv_msg = |Row { <up>-row }: no quantity to scrap for material { lv_mat }| ).
        <up>-err = abap_true.
      ENDIF.
      IF <up>-unres < 0 OR <up>-qi < 0 OR <up>-blk < 0.
        add_log( iv_ok = abap_false iv_row = <up>-row iv_matnr = <up>-matnr
                 iv_msg = |Row { <up>-row }: negative quantity for material { lv_mat }| ).
        <up>-err = abap_true.
      ENDIF.

      " 6. UoM must be the base unit (additional check - sample has PCK rows)
      IF <up>-meins IS NOT INITIAL AND <up>-meins <> <mara>-meins.
        add_log( iv_ok = abap_false iv_row = <up>-row iv_matnr = <up>-matnr
                 iv_msg = |Row { <up>-row }: UOM { <up>-meins } differs from base UOM { <mara>-meins } of material { lv_mat }| ).
        <up>-err = abap_true.
      ENDIF.

      " 7. stock check (MCHB >= upload) - record missing = no stock
      IF <up>-werks IS NOT INITIAL AND <up>-lgort IS NOT INITIAL AND <up>-charg IS NOT INITIAL.
        READ TABLE lt_mchb ASSIGNING FIELD-SYMBOL(<mchb>)
             WITH KEY matnr = <up>-matnr werks = <up>-werks
                      lgort = <up>-lgort charg = <up>-charg.
        IF sy-subrc <> 0
           OR <mchb>-clab  < <up>-unres
           OR <mchb>-cinsm < <up>-qi
           OR <mchb>-cspem < <up>-blk.
          add_log( iv_ok = abap_false iv_row = <up>-row iv_matnr = <up>-matnr
                   iv_werks = <up>-werks iv_lgort = <up>-lgort iv_charg = <up>-charg
                   iv_msg = |Material { lv_mat } { <up>-werks } { <up>-lgort } not in stock| ).
          <up>-err = abap_true.
        ENDIF.
      ENDIF.

      " valuation type for later posting (batch managed materials only)
      IF <up>-xchpf = abap_true.
        READ TABLE lt_mcha ASSIGNING FIELD-SYMBOL(<mcha>)
             WITH KEY matnr = <up>-matnr werks = <up>-werks charg = <up>-charg.
        IF sy-subrc = 0.
          <up>-bwtar = <mcha>-bwtar.
        ENDIF.
      ELSE.
        CLEAR <up>-charg.            " non batch managed: no batch in BAPI
      ENDIF.
    ENDLOOP.

    " 8. duplicates material + plant + sloc
    DATA(lt_sorted) = mt_up.
    SORT lt_sorted BY matnr werks lgort.
    DATA(lv_prev) = VALUE ty_up( ).
    DATA(lv_dup_done) = abap_false.
    LOOP AT lt_sorted ASSIGNING FIELD-SYMBOL(<s>).
      IF sy-tabix > 1 AND <s>-matnr = lv_prev-matnr AND <s>-werks = lv_prev-werks
         AND <s>-lgort = lv_prev-lgort AND <s>-matnr IS NOT INITIAL.
        add_log( iv_ok = abap_false iv_row = <s>-row iv_matnr = <s>-matnr
                 iv_werks = <s>-werks iv_lgort = <s>-lgort
                 iv_msg = |Duplicate { mat_out( <s>-matnr ) } { <s>-werks } { <s>-lgort } found| ).
        MODIFY mt_up FROM VALUE #( err = abap_true ) TRANSPORTING err
               WHERE matnr = <s>-matnr AND werks = <s>-werks AND lgort = <s>-lgort.
      ENDIF.
      lv_prev = <s>.
    ENDLOOP.

    " ---- result ---------------------------------------------------------
    rv_ok = xsdbool( mt_log IS INITIAL ).
  ENDMETHOD.

  "--------------------------------------------------------------------
  " Posting
  "--------------------------------------------------------------------
  METHOD process.
    " one ZSCRAP number per upload
    CALL FUNCTION 'NUMBER_GET_NEXT'
      EXPORTING
        nr_range_nr = gc_nr_range
        object      = gc_nr_object
      IMPORTING
        number      = mv_number
      EXCEPTIONS
        OTHERS      = 1.
    IF sy-subrc <> 0.
      add_log( iv_ok = abap_false
               iv_msg = |No number could be drawn from number range { gc_nr_object }| ).
      RETURN.
    ENDIF.

    LOOP AT mt_up ASSIGNING FIELD-SYMBOL(<up>).
      process_row( <up> ).
    ENDLOOP.
  ENDMETHOD.

  METHOD process_row.
    " Each source stock type is its own sequence; first failure stops the row
    IF is_row-unres > 0.
      IF post_seq( is_row = is_row iv_xfer = '344' iv_rev = '343'
                   iv_qty = is_row-unres ) = abap_false.
        RETURN.
      ENDIF.
    ENDIF.
    IF is_row-qi > 0.
      IF post_seq( is_row = is_row iv_xfer = '350' iv_rev = '349'
                   iv_qty = is_row-qi ) = abap_false.
        RETURN.
      ENDIF.
    ENDIF.
    IF is_row-blk > 0.
      post_seq( is_row = is_row iv_xfer = '' iv_rev = ''
                iv_qty = is_row-blk ).
    ENDIF.
  ENDMETHOD.

  METHOD post_seq.
    " step 1: transfer to blocked (not for blocked stock)
    IF iv_xfer IS NOT INITIAL.
      IF post_gm( is_row = is_row iv_bwart = iv_xfer iv_qty = iv_qty ) = abap_false.
        RETURN.                                   " nothing posted -> next
      ENDIF.
    ENDIF.

    " step 2: scrap
    IF post_gm( is_row = is_row iv_bwart = '555' iv_qty = iv_qty ) = abap_true.
      rv_ok = abap_true.
      RETURN.
    ENDIF.

    " scrap failed after the transfer posted -> put the stock back
    IF iv_xfer IS NOT INITIAL.
      post_gm( is_row = is_row iv_bwart = iv_rev iv_qty = iv_qty
               iv_note = 'Reversal of transfer after failed scrap: ' ).
    ENDIF.
  ENDMETHOD.

  METHOD post_gm.
    DATA: ls_head   TYPE bapi2017_gm_head_01,
          ls_code   TYPE bapi2017_gm_code,
          ls_item   TYPE bapi2017_gm_item_create,
          lt_item   TYPE STANDARD TABLE OF bapi2017_gm_item_create,
          lt_return TYPE STANDARD TABLE OF bapiret2,
          lv_matdoc TYPE bapi2017_gm_head_ret-mat_doc,
          lv_year   TYPE bapi2017_gm_head_ret-doc_year.

    ls_head-pstng_date = sy-datum.
    ls_head-doc_date   = sy-datum.
    ls_head-header_txt = |Spare Disposal SP{ mv_number }|.

    ls_item-material   = is_row-matnr.
    ls_item-plant      = is_row-werks.
    ls_item-stge_loc   = is_row-lgort.
    ls_item-batch      = is_row-charg.        " blank when not batch managed
    ls_item-val_type   = is_row-bwtar.        " MCHA-BWTAR when batch managed
    ls_item-move_type  = iv_bwart.
    ls_item-entry_qnt  = iv_qty.
    ls_item-entry_uom  = is_row-meins.
    ls_item-item_text  = is_row-reason.

    CASE iv_bwart.
      WHEN '344' OR '350' OR '343' OR '349'.
        ls_code-gm_code = '04'.               " transfer posting
        ls_item-move_plant  = is_row-werks.   " from / to
        ls_item-move_stloc  = is_row-lgort.
        ls_item-move_batch  = is_row-charg.
        ls_item-move_val_type = is_row-bwtar.
      WHEN '555'.
        ls_code-gm_code = '03'.               " GI scrapping from blocked stock
        ls_item-costcenter  = gc_costcenter.
        ls_item-move_reas   = gc_reason.
    ENDCASE.

    APPEND ls_item TO lt_item.

    CALL FUNCTION 'BAPI_GOODSMVT_CREATE'
      EXPORTING
        goodsmvt_header  = ls_head
        goodsmvt_code    = ls_code
      IMPORTING
        materialdocument = lv_matdoc
        matdocumentyear  = lv_year
      TABLES
        goodsmvt_item    = lt_item
        return           = lt_return.

    READ TABLE lt_return ASSIGNING FIELD-SYMBOL(<ret>)
         WITH KEY type = 'E'.
    IF sy-subrc <> 0.
      READ TABLE lt_return ASSIGNING <ret> WITH KEY type = 'A'.
    ENDIF.

    IF sy-subrc = 0 OR lv_matdoc IS INITIAL.
      CALL FUNCTION 'BAPI_TRANSACTION_ROLLBACK'.
      add_log( iv_ok = abap_false iv_row = is_row-row
               iv_matnr = is_row-matnr iv_bwart = iv_bwart iv_menge = iv_qty
               iv_werks = is_row-werks iv_lgort = is_row-lgort iv_charg = is_row-charg
               iv_msg = |{ iv_note }{ COND string( WHEN <ret> IS ASSIGNED THEN <ret>-message
                                                   ELSE 'No material document created' ) }| ).
      RETURN.
    ENDIF.

    CALL FUNCTION 'BAPI_TRANSACTION_COMMIT'
      EXPORTING
        wait = abap_true.

    add_log( iv_ok = abap_true iv_row = is_row-row
             iv_matnr = is_row-matnr iv_bwart = iv_bwart iv_menge = iv_qty
             iv_werks = is_row-werks iv_lgort = is_row-lgort iv_charg = is_row-charg
             iv_mblnr = lv_matdoc iv_mjahr = lv_year
             iv_msg = |{ iv_note }Material Documents Posted| ).
    rv_ok = abap_true.
  ENDMETHOD.

  "--------------------------------------------------------------------
  " Attach approval e-mail once, link to every material document created
  "--------------------------------------------------------------------
  METHOD attach_email.
    DATA: ls_folder  TYPE soodk,
          ls_doc     TYPE sodocchgi1,
          ls_info    TYPE sofolenti1,
          lt_head    TYPE STANDARD TABLE OF soli,
          lt_content TYPE solix_tab,
          ls_obja    TYPE borident,
          ls_objb    TYPE borident,
          lv_len     TYPE i.

    DATA(lt_docs) = VALUE tt_log( FOR l IN mt_log WHERE ( status = gc_ok AND mblnr IS NOT INITIAL ) ( l ) ).
    IF lt_docs IS INITIAL OR mt_mail IS INITIAL.
      RETURN.
    ENDIF.

    CALL FUNCTION 'SO_FOLDER_ROOT_ID_GET'
      EXPORTING
        region    = 'B'
      IMPORTING
        folder_id = ls_folder
      EXCEPTIONS
        OTHERS    = 1.
    IF sy-subrc <> 0.
      add_log( iv_ok = abap_false iv_msg = 'Approval email: SAPoffice root folder not found' ).
      RETURN.
    ENDIF.

    " file name (without path) for the attachment
    DATA(lv_name) = segment( val   = replace( val = mv_mail sub = '\' with = '/' occ = 0 )
                             index = -1
                             sep   = '/' ).

    lt_content = mt_mail.
    lv_len     = mv_maillen.

    ls_doc-obj_name  = 'MESSAGE'.
    ls_doc-obj_descr = lv_name.
    ls_doc-obj_langu = sy-langu.
    ls_doc-doc_size  = lv_len.

    APPEND VALUE soli( line = |&SO_FILENAME={ lv_name }| ) TO lt_head.
    APPEND VALUE soli( line = '&SO_FORMAT=BIN' )            TO lt_head.

    CALL FUNCTION 'SO_DOCUMENT_INSERT_API1'
      EXPORTING
        folder_id                  = ls_folder
        document_data              = ls_doc
        document_type              = 'EXT'
      IMPORTING
        document_info              = ls_info
      TABLES
        object_header              = lt_head
        contents_hex               = lt_content
      EXCEPTIONS
        OTHERS                     = 1.
    IF sy-subrc <> 0.
      add_log( iv_ok = abap_false iv_msg = 'Approval email could not be stored (SO_DOCUMENT_INSERT_API1)' ).
      RETURN.
    ENDIF.

    " uploaded once, now associate with every material document
    ls_objb-objkey  = ls_info-doc_id.
    ls_objb-objtype = 'MESSAGE'.
    ls_obja-objtype = 'BUS2017'.
    SORT lt_docs BY mblnr mjahr.
    DELETE ADJACENT DUPLICATES FROM lt_docs COMPARING mblnr mjahr.

    LOOP AT lt_docs ASSIGNING FIELD-SYMBOL(<d>).
      ls_obja-objkey = |{ <d>-mblnr }{ <d>-mjahr }|.
      CALL FUNCTION 'BINARY_RELATION_CREATE'
        EXPORTING
          obj_rolea    = ls_obja
          obj_roleb    = ls_objb
          relationtype = 'ATTA'
        EXCEPTIONS
          OTHERS       = 1.
      IF sy-subrc <> 0.
        add_log( iv_ok = abap_false iv_mblnr = <d>-mblnr iv_mjahr = <d>-mjahr
                 iv_msg = |Approval email could not be linked to material document { <d>-mblnr }| ).
      ENDIF.
    ENDLOOP.

    COMMIT WORK AND WAIT.
  ENDMETHOD.

  "--------------------------------------------------------------------
  " Output
  "--------------------------------------------------------------------
  METHOD build_salv.
    cl_salv_table=>factory( EXPORTING list_display = iv_list
                            IMPORTING r_salv_table = eo_salv
                            CHANGING  t_table      = ct_log ).
    DATA(lo_cols) = eo_salv->get_columns( ).
    lo_cols->set_optimize( abap_true ).
    lo_cols->set_color_column( 'ROWCOL' ).
    TRY.
        lo_cols->get_column( 'STATUS'  )->set_short_text( 'Status' ).
        lo_cols->get_column( 'MBLNR'   )->set_short_text( 'Mat.Doc.' ).
        lo_cols->get_column( 'MATNR'   )->set_short_text( 'Material' ).
        lo_cols->get_column( 'BWART'   )->set_short_text( 'Mvt Type' ).
        lo_cols->get_column( 'MENGE'   )->set_short_text( 'Quantity' ).
        lo_cols->get_column( 'WERKS'   )->set_short_text( 'Plant' ).
        lo_cols->get_column( 'LGORT'   )->set_short_text( 'Sloc' ).
        lo_cols->get_column( 'CHARG'   )->set_short_text( 'Batch' ).
        lo_cols->get_column( 'MESSAGE' )->set_short_text( 'Message' ).
        lo_cols->get_column( 'ROW'     )->set_short_text( 'Xls Row' ).
        lo_cols->get_column( 'MJAHR'   )->set_technical( abap_true ).
        lo_cols->get_column( 'ROWCOL'  )->set_technical( abap_true ).
      CATCH cx_salv_not_found.
    ENDTRY.
    eo_salv->get_functions( )->set_all( abap_true ).
  ENDMETHOD.

  METHOD show_alv.
    IF mt_log IS INITIAL.
      RETURN.
    ENDIF.
    build_salv( EXPORTING iv_list = abap_false
                IMPORTING eo_salv = DATA(lo_salv)
                CHANGING  ct_log  = mt_log ).
    lo_salv->display( ).
  ENDMETHOD.

  METHOD send_mail.
    DATA: lt_body TYPE bcsy_text.

    IF s_email[] IS INITIAL OR mt_log IS INITIAL.
      RETURN.
    ENDIF.

    " results as xlsx
    DATA(lt_copy) = mt_log.
    build_salv( EXPORTING iv_list = abap_true
                IMPORTING eo_salv = DATA(lo_salv)
                CHANGING  ct_log  = lt_copy ).
    DATA(lv_xlsx) = lo_salv->to_xml( xml_type = if_salv_bs_xml=>c_type_xlsx ).
    DATA(lt_att)  = cl_bcs_convert=>xstring_to_solix( iv_xstring = lv_xlsx ).

    APPEND 'Hi Team,' TO lt_body.
    APPEND '' TO lt_body.
    APPEND 'Attached is the result of Spares Scrapping Upload:' TO lt_body.
    APPEND '' TO lt_body.
    APPEND |ZSCRAP number: SP{ mv_number }| TO lt_body.
    APPEND '' TO lt_body.
    APPEND `This mailbox is not monitored, please don't reply. Please raise Optus IT support ticket for any issues.` TO lt_body.

    TRY.
        DATA(lo_bcs) = cl_bcs=>create_persistent( ).
        DATA(lo_doc) = cl_document_bcs=>create_document(
                         i_type    = 'RAW'
                         i_text    = lt_body
                         i_subject = 'Spare Scrapping Upload Result' ).
        lo_doc->add_attachment(
          i_attachment_type    = 'BIN'
          i_attachment_subject = 'Spare_Scrapping_Upload_Result'
          i_attachment_size    = CONV so_obj_len( xstrlen( lv_xlsx ) )
          i_attachment_header  = VALUE #( ( line = '&SO_FILENAME=Spare_Scrapping_Upload_Result.xlsx' ) )
          i_att_content_hex    = lt_att ).
        lo_bcs->set_document( lo_doc ).

        LOOP AT s_email ASSIGNING FIELD-SYMBOL(<m>) WHERE sign = 'I' AND option = 'EQ'.
          lo_bcs->add_recipient(
            i_recipient = cl_cam_address_bcs=>create_internet_address( CONV #( <m>-low ) ) ).
        ENDLOOP.
        lo_bcs->set_send_immediately( abap_true ).
        lo_bcs->send( ).
        COMMIT WORK.
      CATCH cx_bcs INTO DATA(lx).
        MESSAGE |E-mail could not be sent: { lx->get_text( ) }| TYPE 'S' DISPLAY LIKE 'W'.
    ENDTRY.
  ENDMETHOD.

ENDCLASS.

*&---------------------------------------------------------------------*
*& INCLUDE ZMMSPARESCRAP_MAIN
*&---------------------------------------------------------------------*
INITIALIZATION.
  t_b1 = 'Spares Scrap Proposal'.
  %_p_file_%_app_%-text  = 'Upload File'.
  %_p_mail_%_app_%-text  = 'Approval Email File'.
  %_s_email_%_app_%-text = 'Email Result'.
  %_p_test_%_app_%-text  = 'Test Run'.

AT SELECTION-SCREEN ON VALUE-REQUEST FOR p_file.
  lcl_scrap=>f4_file( EXPORTING iv_filter = |Excel (*.xlsx)\|*.xlsx|
                      CHANGING  cv_file   = p_file ).

AT SELECTION-SCREEN ON VALUE-REQUEST FOR p_mail.
  lcl_scrap=>f4_file( EXPORTING iv_filter = |Email (*.msg;*.pdf)\|*.msg;*.pdf|
                      CHANGING  cv_file   = p_mail ).

START-OF-SELECTION.
  NEW lcl_scrap( iv_file = p_file
                 iv_mail = p_mail
                 iv_test = p_test )->run( ).
