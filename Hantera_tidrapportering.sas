/* Conditionally delete set of tables or views, if they exists          */
/* If the member does not exist, then no action is performed   */
%macro _eg_conditional_dropds /parmbuff;
	
   	%local num;
   	%local stepneeded;
   	%local stepstarted;
   	%local dsname;
	%local name;

   	%let num=1;
	/* flags to determine whether a PROC SQL step is needed */
	/* or even started yet                                  */
	%let stepneeded=0;
	%let stepstarted=0;
   	%let dsname= %qscan(&syspbuff,&num,',()');
	%do %while(&dsname ne);	
		%let name = %sysfunc(left(&dsname));
		%if %qsysfunc(exist(&name)) %then %do;
			%let stepneeded=1;
			%if (&stepstarted eq 0) %then %do;
				proc sql;
				%let stepstarted=1;

			%end;
				drop table &name;
		%end;

		%if %sysfunc(exist(&name,view)) %then %do;
			%let stepneeded=1;
			%if (&stepstarted eq 0) %then %do;
				proc sql;
				%let stepstarted=1;
			%end;
				drop view &name;
		%end;
		%let num=%eval(&num+1);
      	%let dsname=%qscan(&syspbuff,&num,',()');
	%end;
	%if &stepstarted %then %do;
		quit;
	%end;
%mend _eg_conditional_dropds;


/* Build where clauses from stored process parameters */
%macro _eg_WhereParam( COLUMN, PARM, OPERATOR, TYPE=S, MATCHALL=_ALL_VALUES_, MATCHALL_CLAUSE=1, MAX= , IS_EXPLICIT=0, MATCH_CASE=1);

  %local q1 q2 sq1 sq2;
  %local isEmpty;
  %local isEqual isNotEqual;
  %local isIn isNotIn;
  %local isString;
  %local isBetween;

  %let isEqual = ("%QUPCASE(&OPERATOR)" = "EQ" OR "&OPERATOR" = "=");
  %let isNotEqual = ("%QUPCASE(&OPERATOR)" = "NE" OR "&OPERATOR" = "<>");
  %let isIn = ("%QUPCASE(&OPERATOR)" = "IN");
  %let isNotIn = ("%QUPCASE(&OPERATOR)" = "NOT IN");
  %let isString = (%QUPCASE(&TYPE) eq S or %QUPCASE(&TYPE) eq STRING );
  %if &isString %then
  %do;
	%if "&MATCH_CASE" eq "0" %then %do;
		%let COLUMN = %str(UPPER%(&COLUMN%));
	%end;
	%let q1=%str(%");
	%let q2=%str(%");
	%let sq1=%str(%'); 
	%let sq2=%str(%'); 
  %end;
  %else %if %QUPCASE(&TYPE) eq D or %QUPCASE(&TYPE) eq DATE %then 
  %do;
    %let q1=%str(%");
    %let q2=%str(%"d);
	%let sq1=%str(%'); 
    %let sq2=%str(%'); 
  %end;
  %else %if %QUPCASE(&TYPE) eq T or %QUPCASE(&TYPE) eq TIME %then
  %do;
    %let q1=%str(%");
    %let q2=%str(%"t);
	%let sq1=%str(%'); 
    %let sq2=%str(%'); 
  %end;
  %else %if %QUPCASE(&TYPE) eq DT or %QUPCASE(&TYPE) eq DATETIME %then
  %do;
    %let q1=%str(%");
    %let q2=%str(%"dt);
	%let sq1=%str(%'); 
    %let sq2=%str(%'); 
  %end;
  %else
  %do;
    %let q1=;
    %let q2=;
	%let sq1=;
    %let sq2=;
  %end;
  
  %if "&PARM" = "" %then %let PARM=&COLUMN;

  %let isBetween = ("%QUPCASE(&OPERATOR)"="BETWEEN" or "%QUPCASE(&OPERATOR)"="NOT BETWEEN");

  %if "&MAX" = "" %then %do;
    %let MAX = &parm._MAX;
    %if &isBetween %then %let PARM = &parm._MIN;
  %end;

  %if not %symexist(&PARM) or (&isBetween and not %symexist(&MAX)) %then %do;
    %if &IS_EXPLICIT=0 %then %do;
		not &MATCHALL_CLAUSE
	%end;
	%else %do;
	    not 1=1
	%end;
  %end;
  %else %if "%qupcase(&&&PARM)" = "%qupcase(&MATCHALL)" %then %do;
    %if &IS_EXPLICIT=0 %then %do;
	    &MATCHALL_CLAUSE
	%end;
	%else %do;
	    1=1
	%end;	
  %end;
  %else %if (not %symexist(&PARM._count)) or &isBetween %then %do;
    %let isEmpty = ("&&&PARM" = "");
    %if (&isEqual AND &isEmpty AND &isString) %then
       &COLUMN is null;
    %else %if (&isNotEqual AND &isEmpty AND &isString) %then
       &COLUMN is not null;
    %else %do;
	   %if &IS_EXPLICIT=0 %then %do;
           &COLUMN &OPERATOR 
			%if "&MATCH_CASE" eq "0" %then %do;
				%unquote(&q1)%QUPCASE(&&&PARM)%unquote(&q2)
			%end;
			%else %do;
				%unquote(&q1)&&&PARM%unquote(&q2)
			%end;
	   %end;
	   %else %do;
	       &COLUMN &OPERATOR 
			%if "&MATCH_CASE" eq "0" %then %do;
				%unquote(%nrstr(&sq1))%QUPCASE(&&&PARM)%unquote(%nrstr(&sq2))
			%end;
			%else %do;
				%unquote(%nrstr(&sq1))&&&PARM%unquote(%nrstr(&sq2))
			%end;
	   %end;
       %if &isBetween %then 
          AND %unquote(&q1)&&&MAX%unquote(&q2);
    %end;
  %end;
  %else 
  %do;
	%local emptyList;
  	%let emptyList = %symexist(&PARM._count);
  	%if &emptyList %then %let emptyList = &&&PARM._count = 0;
	%if (&emptyList) %then
	%do;
		%if (&isNotin) %then
		   1;
		%else
			0;
	%end;
	%else %if (&&&PARM._count = 1) %then 
    %do;
      %let isEmpty = ("&&&PARM" = "");
      %if (&isIn AND &isEmpty AND &isString) %then
        &COLUMN is null;
      %else %if (&isNotin AND &isEmpty AND &isString) %then
        &COLUMN is not null;
      %else %do;
	    %if &IS_EXPLICIT=0 %then %do;
			%if "&MATCH_CASE" eq "0" %then %do;
				&COLUMN &OPERATOR (%unquote(&q1)%QUPCASE(&&&PARM)%unquote(&q2))
			%end;
			%else %do;
				&COLUMN &OPERATOR (%unquote(&q1)&&&PARM%unquote(&q2))
			%end;
	    %end;
		%else %do;
		    &COLUMN &OPERATOR (
			%if "&MATCH_CASE" eq "0" %then %do;
				%unquote(%nrstr(&sq1))%QUPCASE(&&&PARM)%unquote(%nrstr(&sq2)))
			%end;
			%else %do;
				%unquote(%nrstr(&sq1))&&&PARM%unquote(%nrstr(&sq2)))
			%end;
		%end;
	  %end;
    %end;
    %else 
    %do;
       %local addIsNull addIsNotNull addComma;
       %let addIsNull = %eval(0);
       %let addIsNotNull = %eval(0);
       %let addComma = %eval(0);
       (&COLUMN &OPERATOR ( 
       %do i=1 %to &&&PARM._count; 
          %let isEmpty = ("&&&PARM&i" = "");
          %if (&isString AND &isEmpty AND (&isIn OR &isNotIn)) %then
          %do;
             %if (&isIn) %then %let addIsNull = 1;
             %else %let addIsNotNull = 1;
          %end;
          %else
          %do;		     
            %if &addComma %then %do;,%end;
			%if &IS_EXPLICIT=0 %then %do;
				%if "&MATCH_CASE" eq "0" %then %do;
					%unquote(&q1)%QUPCASE(&&&PARM&i)%unquote(&q2)
				%end;
				%else %do;
					%unquote(&q1)&&&PARM&i%unquote(&q2)
				%end;
			%end;
			%else %do;
				%if "&MATCH_CASE" eq "0" %then %do;
					%unquote(%nrstr(&sq1))%QUPCASE(&&&PARM&i)%unquote(%nrstr(&sq2))
				%end;
				%else %do;
					%unquote(%nrstr(&sq1))&&&PARM&i%unquote(%nrstr(&sq2))
				%end; 
			%end;
            %let addComma = %eval(1);
          %end;
       %end;) 
       %if &addIsNull %then OR &COLUMN is null;
       %else %if &addIsNotNull %then AND &COLUMN is not null;
       %do;)
       %end;
    %end;
  %end;
%mend _eg_WhereParam;


/* ---------------------------------- */
/* MACRO: enterpriseguide             */
/* PURPOSE: define a macro variable   */
/*   that contains the file system    */
/*   path of the WORK library on the  */
/*   server.  Note that different     */
/*   logic is needed depending on the */
/*   server type.                     */
/* ---------------------------------- */
%macro enterpriseguide;
%global sasworklocation;
%local tempdsn unique_dsn path;

%if &sysscp=OS %then %do; /* MVS Server */
	%if %sysfunc(getoption(filesystem))=MVS %then %do;
        /* By default, physical file name will be considered a classic MVS data set. */
	    /* Construct dsn that will be unique for each concurrent session under a particular account: */
		filename egtemp '&egtemp' disp=(new,delete); /* create a temporary data set */
 		%let tempdsn=%sysfunc(pathname(egtemp)); /* get dsn */
		filename egtemp clear; /* get rid of data set - we only wanted its name */
		%let unique_dsn=".EGTEMP.%substr(&tempdsn, 1, 16).PDSE"; 
		filename egtmpdir &unique_dsn
			disp=(new,delete,delete) space=(cyl,(5,5,50))
			dsorg=po dsntype=library recfm=vb
			lrecl=8000 blksize=8004 ;
		options fileext=ignore ;
	%end; 
 	%else %do; 
        /* 
		By default, physical file name will be considered an HFS 
		(hierarchical file system) file. 
		*/
		%if "%sysfunc(getoption(filetempdir))"="" %then %do;
			filename egtmpdir '/tmp';
		%end;
		%else %do;
			filename egtmpdir "%sysfunc(getoption(filetempdir))";
		%end;
	%end; 
	%let path=%sysfunc(pathname(egtmpdir));
    %let sasworklocation=%sysfunc(quote(&path));  
%end; /* MVS Server */
%else %do;
	%let sasworklocation = "%sysfunc(getoption(work))/";
%end;
%if &sysscp=VMS_AXP %then %do; /* Alpha VMS server */
	%let sasworklocation = "%sysfunc(getoption(work))";                         
%end;
%if &sysscp=CMS %then %do; 
	%let path = %sysfunc(getoption(work));                         
	%let sasworklocation = "%substr(&path, %index(&path,%str( )))";
%end;
%mend enterpriseguide;

%enterpriseguide


/* save the current settings of XPIXELS and YPIXELS */
/* so that they can be restored later               */
%macro _sas_pushchartsize(new_xsize, new_ysize);
	%global _savedxpixels _savedypixels;
	options nonotes;
	proc sql noprint;
	select setting into :_savedxpixels
	from sashelp.vgopt
	where optname eq "XPIXELS";
	select setting into :_savedypixels
	from sashelp.vgopt
	where optname eq "YPIXELS";
	quit;
	options notes;
	GOPTIONS XPIXELS=&new_xsize YPIXELS=&new_ysize;
%mend _sas_pushchartsize;

/* restore the previous values for XPIXELS and YPIXELS */
%macro _sas_popchartsize;
	%if %symexist(_savedxpixels) %then %do;
		GOPTIONS XPIXELS=&_savedxpixels YPIXELS=&_savedypixels;
		%symdel _savedxpixels / nowarn;
		%symdel _savedypixels / nowarn;
	%end;
%mend _sas_popchartsize;


%*--------------------------------------------------------------*
 * Tests the current version against a required version. A      *
 * negative result means that the SAS server version is less    *
 * than the version required.  A positive result means that     *
 * the SAS server version is greater than the version required. *
 * A result of zero indicates that the SAS server is exactly    *
 * the version required.                                        *
 *                                                              *
 * NOTE: The parameter maint is optional.                       *
 *--------------------------------------------------------------*;
%macro _SAS_VERCOMP(major, minor, maint);
    %_SAS_VERCOMP_FV(&major, &minor, &maint, &major, &minor, &maint)
%mend _SAS_VERCOMP;

%*--------------------------------------------------------------*
 * Tests the current version against either the required        *
 * foundation or Viya required version depending on whether the *
 * SYSVLONG version is a foundation or Viya one. A negative     *
 * result means that the SAS server version is less than the    *
 * version required.  A positive result means that the SAS      *
 * server version is greater than the version required. A       *
 * result of zero indicates that the SAS server is exactly the  *
 * version required.                                            *
 *                                                              *
 * NOTE: The *maint parameters are optional.                    *
 *--------------------------------------------------------------*;
%macro _SAS_VERCOMP_FV(fmajor, fminor, fmaint, vmajor, vminor, vmaint);
    %local major;
    %local minor;
    %local maint;
    %local CurMaj;
    %local CurMin;
    %local CurMnt;

    %* Pull the current version string apart.;
    %let CurMaj = %scan(&sysvlong, 1, %str(.));

    %* The Viya version number has a V on the front which means
       we need to adjust the Maint SCAN funtion index and also
       get the appropriate parameters for the major, minor, and
       maint values we need to check against (foundation or Viya);
    %if %eval(&CurMaj EQ V) %then
        %do;
		   %*   MM mm t           MM = Major version , mm = Minor version , t = Maint version ;
		   %* V.03.04M2P07112018 ;

            %let major = &vmajor;
            %let minor = &vminor;
            %let maint = &vmaint;
			%let CurMaj = %scan(&sysvlong, 2, %str(.));
			%* Index is purposely 2 because V is now one of the scan delimiters ;
			%let CurMin = %scan(&sysvlong, 2, %str(.ABCDEFGHIKLMNOPQRSTUVWXYZ));
			%let CurMnt = %scan(&sysvlong, 3, %str(.ABCDEFGHIKLMNOPQRSTUVWXYZ));
        %end;
    %else
        %do;
		    %* M mm    t           M = Major version , mm = Minor version , t = Maint version ;  
		    %* 9.01.02M0P11212005 ;

            %let major = &fmajor;
            %let minor = &fminor;
            %let maint = &fmaint;
			%let CurMin = %scan(&sysvlong, 2, %str(.));
			%let CurMnt = %scan(&sysvlong, 4, %str(.ABCDEFGHIKLMNOPQRSTUVWXYZ));
        %end;

    %* Now perform the version comparison.;
    %if %eval(&major NE &CurMaj) %then
        %eval(&CurMaj - &major);
    %else
        %if %eval(&minor NE &CurMin) %then
            %eval(&CurMin - &minor);
        %else
            %if "&maint" = "" %then
                %str(0);
            %else
                %eval(&CurMnt - &maint);
%mend _SAS_VERCOMP_FV;

%*--------------------------------------------------------------*
 * This macro calls _SAS_VERCONDCODE_FV() with the passed       *
 * version. If the current server version matches or is newer,  *
 * then the true code (tcode) is executed, else the false code  *
 * (fcode) is executed.                                         *
 * Example:                                                     *
 *  %let isV92 =                                                *
 *     %_SAS_VERCONDCODE(9,2,0,                                 *
 *         tcode=%nrstr(Yes),                                   *
 *         fcode=%nrstr(No))                                    *
 *--------------------------------------------------------------*;
%macro _SAS_VERCONDCODE( major, minor, maint, tcode=, fcode= );
    %_SAS_VERCONDCODE_FV( &major, &minor, &maint, &major, &minor, &maint, &tcode, fcode )
%mend _SAS_VERCONDCODE;

%*--------------------------------------------------------------*
 * This macro calls _SAS_VERCOMP_FV() with the passed versions. *
 * If the current server version matches or is newer, then the  *
 * true code (tcode) is executed, else the false code (fcode)   *
 * is executed.                                                 *
 * Example:                                                     *
 *  %let isV92 =                                                *
 *     %_SAS_VERCONDCODE_FV(9,2,0, 3,5,0                        *
 *         tcode=%nrstr(Yes),                                   *
 *         fcode=%nrstr(No))                                    *
 *--------------------------------------------------------------*;
%macro _SAS_VERCONDCODE_FV( fmajor, fminor, fmaint, vmajor, vminor, vmaint, tcode=, fcode= );
    %if %_SAS_VERCOMP_FV(&fmajor, &fminor, &fmaint, &vmajor, &vminor, &vmaint) >= 0 %then
        %do;
        &tcode
        %end;
    %else
        %do;
        &fcode
        %end;
%mend _SAS_VERCONDCODE_FV;

%*--------------------------------------------------------------*
 * Tests the current version to see if it is a Viya version     *
 * number.                                                      *
 * A result of 1 indicates that the SAS server is a Viya        *
 * server.                                                      *
 * A zero result indicates that the server version is not       *
 * that of a Viya server.                                       *
 *--------------------------------------------------------------*;
%macro _SAS_ISVIYA;
    %local Major;

    %* Get the major component of the current version string.;
    %let Major = %scan(&sysvlong, 1, %str(.));

    %* Check if it it V for Viya.;
    %if %eval(&Major EQ V) %then
        %str(1);
    %else
        %str(0);
%mend _SAS_ISVIYA;


ODS PROCTITLE;
OPTIONS DEV=PNG;
GOPTIONS XPIXELS=0 YPIXELS=0;
%macro HTML5AccessibleGraphSupported;
    %if %_SAS_VERCOMP_FV(9,4,4, 0,0,0) >= 0 %then ACCESSIBLE_GRAPH;
%mend;


/* --------------------------------------------------------------------
   Code generated by a SAS task
   
   Generated on den 5 oktober 2018 at 16:43:26
   By task:     Import Data Wizard
   
   Source file:
   C:\Users\freha296\Documents\Administrativt\All_my_time.tsv
   Server:      Local File System
   
   Output data: WORK.All_my_time
   Server:      Local
   -------------------------------------------------------------------- */

DATA WORK.All_my_time;
    LENGTH
        Work_Item        $ 256
        Start              8
        End                8
        Duration           8
        Notes            $ 256 ;
    LABEL
        Work_Item        = "Work Item" ;
    FORMAT
        Work_Item        $CHAR256.
        Start            DATETIME18.
        End              DATETIME18.
        Duration         TIME8.
        Notes            $CHAR256. ;
    INFORMAT
        Work_Item        $CHAR256.
        Start            DATETIME18.
        End              DATETIME18.
        Duration         TIME11.
        Notes            $CHAR256. ;
    INFILE 'C:\Users\freha296\Documents\Administrativt\All_my_time.tsv'
        LRECL=32767
        FIRSTOBS=2
        ENCODING="UTF-8"
        DLM='2c'x
        MISSOVER
        DSD ;
    INPUT
        Work_Item        : $CHAR256.
        Start            : ?? ANYDTDTM19.
        End              : ?? ANYDTDTM19.
        Duration         : ?? TIME11.
        Notes            : $CHAR256. ;
RUN;




%LET _CLIENTTASKLABEL=;
%LET _CLIENTPROCESSFLOWNAME=;
%LET _CLIENTPROJECTPATH=;
%LET _CLIENTPROJECTPATHHOST=;
%LET _CLIENTPROJECTNAME=;
%LET _SASPROGRAMFILE=;
%LET _SASPROGRAMFILEHOST=;


%_eg_conditional_dropds(WORK.DERIVED_COLUMNS);

PROC SQL;
   CREATE TABLE WORK.DERIVED_COLUMNS(label="All tid plus deriverade kolumner") AS 
   SELECT t1.Work_Item, 
          t1.Start, 
          t1.End, 
          t1.Duration, 
          /* startdatum */
            (datepart(t1.Start)) FORMAT=yymmdd10. LENGTH=8 AS startdatum, 
          /* manad */
            (month(datepart(t1.Start))) FORMAT=8. LENGTH=8 LABEL="Månad" AS manad, 
          t1.Notes
      FROM WORK.ALL_MY_TIME t1;
QUIT;



%LET _CLIENTTASKLABEL=;
%LET _CLIENTPROCESSFLOWNAME=;
%LET _CLIENTPROJECTPATH=;
%LET _CLIENTPROJECTPATHHOST=;
%LET _CLIENTPROJECTNAME=;




* Kontrollera att filen endast innehåller en månad	;
Proc SQL noprint;
	select min(startdatum), max(startdatum) into :min_startdatum, :max_startdatum
	from work.derived_columns;
quit;

%let antal_rader=&SQLOBS;

%put &=antal_rader;



%macro varna_om_antal_manader;
	%if &antal_rader NE 1 %then %do;
		%put ERROR: &antal_rader rader hittades i work.derived_columns. ;
		%put ERROR: Sessionen avbryts;
		%abort abend; 
	%end;
%mend varna_om_antal_manader;

%varna_om_antal_manader




* Skapar upp alla dagar	för att fylla ut datasetet (för t.ex. lediga dagar)	;
%put &min_startdatum &max_startdatum;

data alla_dagar;
	set work.derived_columns(obs=1 keep=startdatum duration);
	forsta=intnx('MONTH', &min_startdatum, 0, 'BEGINNING');
	sista=intnx('MONTH', &max_startdatum, 0, 'END');
	duration=0;
	do startdatum= forsta to sista;
		output;
	end;
	drop forsta sista;
run;

%LET _CLIENTTASKLABEL=;
%LET _CLIENTPROCESSFLOWNAME=;
%LET _CLIENTPROJECTPATH=;
%LET _CLIENTPROJECTPATHHOST=;
%LET _CLIENTPROJECTNAME=;
%LET _SASPROGRAMFILE=;
%LET _SASPROGRAMFILEHOST=;


%_eg_conditional_dropds(WORK.QUERY_FOR_THISMONTH_DERIVED_COLU);

PROC SQL;
   CREATE TABLE WORK.QUERY_FOR_THISMONTH_DERIVED_COLU AS 
   SELECT t1.Work_Item, 
          t1.Start, 
          t1.End, 
          /* Duration */
            (coalesce(t1.Duration,t2.Duration)) FORMAT=HHMM9. LABEL="Arbetad tid" AS Duration, 
          /* DurationTimmar */
            (CASE  
               WHEN (coalesce(t1.Duration,t2.Duration)) = .
               THEN .
               ELSE (coalesce(t1.Duration,t2.Duration))/3600
            END) FORMAT=COMMA10.2 LENGTH=8 LABEL="Arbetad tid i timmar" AS DurationTimmar, 
          /* startdatum */
            (COALESCE(t1.startdatum, t2.startdatum)) FORMAT=yymmdd10. AS startdatum, 
          /* Vecka */
            (week(COALESCE(t1.startdatum, t2.startdatum), 'V')) AS Vecka, 
          /* veckodag */
            (put(COALESCE(t1.startdatum, t2.startdatum), weekdate3.)) AS veckodag, 
          /* manad */
            ((month((COALESCE(t1.startdatum, t2.startdatum))))) LENGTH=3 LABEL="Månad" AS manad, 
          /* ar */
            (year(COALESCE(t1.startdatum, t2.startdatum))) LENGTH=3 LABEL="År" AS ar, 
          t1.Notes
      FROM WORK.DERIVED_COLUMNS t1
           FULL JOIN WORK.ALLA_DAGAR t2 ON (t1.startdatum = t2.startdatum);
QUIT;




%LET _CLIENTTASKLABEL=;
%LET _CLIENTPROCESSFLOWNAME=;
%LET _CLIENTPROJECTPATH=;
%LET _CLIENTPROJECTPATHHOST=;
%LET _CLIENTPROJECTNAME=;


%LET _CLIENTTASKFILTER = startdatum >= INTNX('MONTH', today(), -1, 'B');
TITLE;
TITLE1 "Tabeller för Antura";
FOOTNOTE;
FOOTNOTE1 "Tid i timmar och decimaler (inte minuter!)";
/* -------------------------------------------------------------------
   Code generated by SAS Task

   Generated on: Thursday, December 4, 2025 at 4:33:40 PM
   By task: Underlag Antura

   Input Data: Local:WORK.QUERY_FOR_THISMONTH_DERIVED_COLU
   Server:  Local
   ------------------------------------------------------------------- */


/* -------------------------------------------------------------------
   Run the tabulate procedure
   ------------------------------------------------------------------- */
PROC TABULATE
DATA=WORK.QUERY_FOR_THISMONTH_DERIVED_COLU(FIRSTOBS=1 )
	FORMAT=COMMA8.2
	;
	WHERE( startdatum >= INTNX('MONTH', today(), -1, 'B'));
	
	VAR DurationTimmar Duration;
	CLASS Work_Item /	ORDER=UNFORMATTED MISSING;
	CLASS veckodag /	ORDER=UNFORMATTED MISSING;
	CLASS startdatum /	ORDER=UNFORMATTED MISSING;
	CLASS Vecka /	ORDER=UNFORMATTED MISSING;
	TABLE /* Page Dimension */
Vecka,
/* Row Dimension */
Work_Item={LABEL="" STYLE(CLASSLEV)={JUST=RIGHT}}*
  DurationTimmar={LABEL=""} 
DurationTimmar={LABEL="Dagsumma"},
/* Column Dimension */
startdatum={LABEL=""}*
  veckodag={LABEL=""}*
    Sum={LABEL=""}*F=COMMA10.2 
ALL={LABEL="Veckosumma"}*
  Sum*F=COMMA10.2
/*Table Options */
/ MISSTEXT=" " 		;
	;

RUN;
/* -------------------------------------------------------------------
   End of task code
   ------------------------------------------------------------------- */
RUN; QUIT;
TITLE; FOOTNOTE;
%SYMDEL _CLIENTTASKFILTER;


%LET _CLIENTTASKLABEL=;
%LET _CLIENTPROCESSFLOWNAME=;
%LET _CLIENTPROJECTPATH=;
%LET _CLIENTPROJECTPATHHOST=;
%LET _CLIENTPROJECTNAME=;


%_eg_conditional_dropds(WORK.typ_av_tid);

PROC SQL;
   CREATE TABLE WORK.typ_av_tid AS 
   SELECT t1.Work_Item, 
          t1.Start, 
          t1.End, 
          t1.Duration, 
          t1.DurationTimmar, 
          t1.startdatum, 
          t1.Vecka, 
          t1.veckodag, 
          t1.manad, 
          t1.ar, 
          /* Typ_av_tid */
            (CASE 
               WHEN t1.Work_Item NOT LIKE 'Från%' THEN 'Arbetstid'
               ELSE t1.Work_Item
            END) FORMAT=$CHAR256. LABEL="Typ av tid" AS Typ_av_tid
      FROM WORK.QUERY_FOR_THISMONTH_DERIVED_COLU t1
      WHERE t1.startdatum > %sysfunc(intnx(MONTH, %sysfunc(today()), -12, END));
QUIT;



%LET _CLIENTTASKLABEL=;
%LET _CLIENTPROCESSFLOWNAME=;
%LET _CLIENTPROJECTPATH=;
%LET _CLIENTPROJECTPATHHOST=;
%LET _CLIENTPROJECTNAME=;


TITLE;
TITLE1 "Till Flex-rapport";
FOOTNOTE;
/* -------------------------------------------------------------------
   Code generated by SAS Task

   Generated on: Thursday, December 4, 2025 at 4:33:40 PM
   By task: Underlag - Flextidsrapport

   Input Data: Local:WORK.TYP_AV_TID
   Server:  Local
   ------------------------------------------------------------------- */


/* -------------------------------------------------------------------
   Run the tabulate procedure
   ------------------------------------------------------------------- */
PROC TABULATE
DATA=WORK.TYP_AV_TID(FIRSTOBS=1 )
	FORMAT=HHMM5.
	;
	
	VAR Duration;
	CLASS Typ_av_tid /	ORDER=UNFORMATTED MISSING;
	CLASS veckodag /	ORDER=UNFORMATTED MISSING;
	CLASS startdatum /	ORDER=UNFORMATTED MISSING;
	CLASS manad /	ORDER=UNFORMATTED MISSING;
	TABLE /* Page Dimension */
manad,
/* Row Dimension */
startdatum*
  veckodag 
ALL={LABEL="Total (ALL)"},
/* Column Dimension */
Typ_av_tid*
  Duration={LABEL=""}*
    Sum={LABEL=""} 
ALL={LABEL="Total (ALL)"}*
  Duration={LABEL=""}*
    Sum={LABEL=""}
/*Table Options */
/ MISSTEXT="" INDENT=0 		;
	;

RUN;
/* -------------------------------------------------------------------
   End of task code
   ------------------------------------------------------------------- */
RUN; QUIT;
TITLE; FOOTNOTE;


%LET _CLIENTTASKLABEL=;
%LET _CLIENTPROCESSFLOWNAME=;
%LET _CLIENTPROJECTPATH=;
%LET _CLIENTPROJECTPATHHOST=;
%LET _CLIENTPROJECTNAME=;


data _null_;
	call symputx('first_day_last_month', INTNX('MONTH', today(), -1, 'B'));
	call symputx('today', today());
run;


proc SQL noprint;
	select distinct ar, vecka into :ar1-, :vecka1-
	from WORK.QUERY_FOR_THISMONTH_DERIVED_COLU
	WHERE startdatum GE &first_day_last_month and startdatum LE &today
	;
quit;
%let antal_veckor=&SQLOBS;


TITLE;
TITLE1 "Tabeller för Antura";
TITLE2 h=1 "Tid i timmar och decimaler (inte minuter!)";
FOOTNOTE;

ods layout gridded columns=2 advance=bygroup style={background=lightgray};

%macro loop_weeks(antal_veckor=);
	%local i where;
	%do i = 1 %to &antal_veckor;

		%let where = ar=&&ar&i and vecka=&&vecka&i;

		title "Vecka &&vecka&i";
		ods region;

		/* -------------------------------------------------------------------
		   Run the tabulate procedure
		   ------------------------------------------------------------------- */
		PROC TABULATE
		DATA=WORK.QUERY_FOR_THISMONTH_DERIVED_COLU FORMAT=COMMA8.2;
			WHERE( &where );
			* by vecka;
			VAR DurationTimmar Duration;
			CLASS Work_Item /	ORDER=UNFORMATTED MISSING;
			CLASS veckodag /	ORDER=UNFORMATTED MISSING;
			CLASS startdatum /	ORDER=UNFORMATTED MISSING;
			TABLE /* Page Dimension */
		/* Row Dimension */
		Work_Item*DurationTimmar={LABEL=""} 
		DurationTimmar={LABEL="Dagsumma"},
		/* Column Dimension */
		startdatum*veckodag={LABEL=""}*Sum={LABEL=""}*F=COMMA10.2 
		ALL={LABEL="Veckosumma"}*Sum*F=COMMA10.2
		/*Table Options */
		/ MISSTEXT=" " 		;
			;

		RUN;

		title "Anteckningar";
		ods region;

		proc SQL;
			select startdatum, notes, durationTimmar
			from WORK.QUERY_FOR_THISMONTH_DERIVED_COLU
			WHERE &where
			and notes NE " "
			;
		quit;

	%end;

%mend loop_weeks;

%loop_weeks(antal_veckor=&antal_veckor)

ods layout end;






/* -------------------------------------------------------------------


options nobyline;

%LET _CLIENTTASKFILTER = startdatum >= INTNX('MONTH', today(), -1, 'B');
TITLE;
TITLE1 "Tabeller för Antura";
FOOTNOTE;
FOOTNOTE1 "Tid i timmar och decimaler (inte minuter!)";


ods layout gridded columns=2 advance=bygroup style={background=lightgray};


PROC TABULATE
DATA=WORK.QUERY_FOR_THISMONTH_DERIVED_COLU
	FORMAT=COMMA8.2
	;
	WHERE( startdatum >= INTNX('MONTH', today(), -1, 'B'));
	by vecka;
	VAR DurationTimmar Duration;
	CLASS Work_Item /	ORDER=UNFORMATTED MISSING;
	CLASS veckodag /	ORDER=UNFORMATTED MISSING;
	CLASS startdatum /	ORDER=UNFORMATTED MISSING;
	CLASS Vecka /	ORDER=UNFORMATTED MISSING;
	TABLE Vecka,
Work_Item*
  DurationTimmar={LABEL=""} 
DurationTimmar={LABEL="Dagsumma"},
startdatum*veckodag={LABEL=""}*Sum={LABEL=""}*F=COMMA10.2 
ALL={LABEL="Veckosumma"}*
  Sum*F=COMMA10.2
/ MISSTEXT=" " 		;
	;

RUN;

TITLE; FOOTNOTE;
%SYMDEL _CLIENTTASKFILTER;

ods layout end ;


   ------------------------------------------------------------------- */

%LET _CLIENTTASKLABEL=;
%LET _CLIENTPROCESSFLOWNAME=;
%LET _CLIENTPROJECTPATH=;
%LET _CLIENTPROJECTPATHHOST=;
%LET _CLIENTPROJECTNAME=;
%LET _SASPROGRAMFILE=;
%LET _SASPROGRAMFILEHOST=;

;*';*";*/;quit;run;
ODS _ALL_ CLOSE;
