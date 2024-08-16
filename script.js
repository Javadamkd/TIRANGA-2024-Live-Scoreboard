{\rtf1\ansi\ansicpg1252\cocoartf2758
\cocoatextscaling0\cocoaplatform0{\fonttbl\f0\fnil\fcharset0 Menlo-Regular;}
{\colortbl;\red255\green255\blue255;\red20\green67\blue174;\red246\green247\blue249;\red46\green49\blue51;
\red24\green25\blue27;\red186\green6\blue115;\red162\green0\blue16;\red77\green80\blue85;\red18\green115\blue126;
}
{\*\expandedcolortbl;;\cssrgb\c9412\c35294\c73725;\cssrgb\c97255\c97647\c98039;\cssrgb\c23529\c25098\c26275;
\cssrgb\c12549\c12941\c14118;\cssrgb\c78824\c15294\c52549;\cssrgb\c70196\c7843\c7059;\cssrgb\c37255\c38824\c40784;\cssrgb\c3529\c52157\c56863;
}
\paperw11900\paperh16840\margl1440\margr1440\vieww11520\viewh8400\viewkind0
\deftab720
\pard\pardeftab720\partightenfactor0

\f0\fs26 \cf2 \cb3 \expnd0\expndtw0\kerning0
\outl0\strokewidth0 \strokec2 function\cf4 \strokec4  \cf5 \strokec5 doGet\cf4 \strokec4 () \{\cb1 \
\pard\pardeftab720\partightenfactor0
\cf4 \cb3   \cf2 \strokec2 return\cf4 \strokec4  \cf6 \cb3 \strokec6 HtmlService\cf4 \cb3 \strokec4 .\cf5 \strokec5 createHtmlOutputFromFile\cf4 \strokec4 (\cf7 \strokec7 'index'\cf4 \strokec4 )\cb1 \
\cb3     .\cf5 \strokec5 setTitle\cf4 \strokec4 (\cf7 \strokec7 'TIRANGA 2024 Live Scoreboard'\cf4 \strokec4 )\cb1 \
\cb3     .\cf5 \strokec5 setXFrameOptionsMode\cf4 \strokec4 (\cf6 \cb3 \strokec6 HtmlService\cf4 \cb3 \strokec4 .\cf6 \cb3 \strokec6 XFrameOptionsMode\cf4 \cb3 \strokec4 .\cf6 \cb3 \strokec6 ALLOWALL\cf4 \cb3 \strokec4 );\cb1 \
\cb3 \}\cb1 \
\
\pard\pardeftab720\partightenfactor0
\cf2 \cb3 \strokec2 function\cf4 \strokec4  \cf5 \strokec5 loadData\cf4 \strokec4 () \{\cb1 \
\pard\pardeftab720\partightenfactor0
\cf4 \cb3   \cf2 \strokec2 const\cf4 \strokec4  \cf5 \strokec5 sheet\cf4 \strokec4  = \cf6 \cb3 \strokec6 SpreadsheetApp\cf4 \cb3 \strokec4 .\cf5 \strokec5 getActiveSpreadsheet\cf4 \strokec4 ().\cf5 \strokec5 getSheetByName\cf4 \strokec4 (\cf7 \strokec7 'Sheet1'\cf4 \strokec4 );\cb1 \
\cb3   \cf2 \strokec2 const\cf4 \strokec4  \cf5 \strokec5 lastRow\cf4 \strokec4  = \cf5 \strokec5 sheet\cf4 \strokec4 .\cf5 \strokec5 getLastRow\cf4 \strokec4 (); \cf8 \strokec8 // Get the last row with data\cf4 \cb1 \strokec4 \
\cb3   \cf2 \strokec2 const\cf4 \strokec4  \cf5 \strokec5 data\cf4 \strokec4  = \cf5 \strokec5 sheet\cf4 \strokec4 .\cf5 \strokec5 getRange\cf4 \strokec4 (\cf9 \strokec9 3\cf4 \strokec4 , \cf9 \strokec9 1\cf4 \strokec4 , \cf5 \strokec5 lastRow\cf4 \strokec4  - \cf9 \strokec9 2\cf4 \strokec4 , \cf9 \strokec9 6\cf4 \strokec4 ).\cf5 \strokec5 getValues\cf4 \strokec4 (); \cf8 \strokec8 // Adjust range to include all rows from row 3\cf4 \cb1 \strokec4 \
\
\cb3   \cf2 \strokec2 let\cf4 \strokec4  \cf5 \strokec5 htmlTable\cf4 \strokec4  = \cf7 \strokec7 ''\cf4 \strokec4 ;\cb1 \
\
\cb3   \cf5 \strokec5 data\cf4 \strokec4 .\cf5 \strokec5 forEach\cf4 \strokec4 (\cf5 \strokec5 row\cf4 \strokec4  => \{\cb1 \
\cb3     \cf5 \strokec5 htmlTable\cf4 \strokec4  += \cf7 \strokec7 '<tr>'\cf4 \strokec4 ;\cb1 \
\cb3     \cf5 \strokec5 row\cf4 \strokec4 .\cf5 \strokec5 forEach\cf4 \strokec4 ((\cf5 \strokec5 cell\cf4 \strokec4 , \cf5 \strokec5 index\cf4 \strokec4 ) => \{\cb1 \
\cb3       \cf2 \strokec2 if\cf4 \strokec4  (\cf5 \strokec5 index\cf4 \strokec4  === \cf9 \strokec9 0\cf4 \strokec4 ) \{\cb1 \
\cb3         \cf5 \strokec5 htmlTable\cf4 \strokec4  += \cf7 \strokec7 `<td class="school-name">\cf4 \strokec4 $\{\cf5 \strokec5 cell\cf4 \strokec4 \}\cf7 \strokec7 </td>`\cf4 \strokec4 ; \cf8 \strokec8 // School name column\cf4 \cb1 \strokec4 \
\cb3       \} \cf2 \strokec2 else\cf4 \strokec4  \cf2 \strokec2 if\cf4 \strokec4  (\cf5 \strokec5 index\cf4 \strokec4  >= \cf9 \strokec9 1\cf4 \strokec4  && \cf5 \strokec5 index\cf4 \strokec4  <= \cf9 \strokec9 4\cf4 \strokec4 ) \{\cb1 \
\cb3         \cf5 \strokec5 htmlTable\cf4 \strokec4  += \cf7 \strokec7 `<td class="marks">\cf4 \strokec4 $\{\cf5 \strokec5 cell\cf4 \strokec4 \}\cf7 \strokec7 </td>`\cf4 \strokec4 ; \cf8 \strokec8 // Round 1 to 4 columns\cf4 \cb1 \strokec4 \
\cb3       \} \cf2 \strokec2 else\cf4 \strokec4  \cf2 \strokec2 if\cf4 \strokec4  (\cf5 \strokec5 index\cf4 \strokec4  === \cf9 \strokec9 5\cf4 \strokec4 ) \{\cb1 \
\cb3         \cf5 \strokec5 htmlTable\cf4 \strokec4  += \cf7 \strokec7 `<td class="total">\cf4 \strokec4 $\{\cf5 \strokec5 cell\cf4 \strokec4 \}\cf7 \strokec7 </td>`\cf4 \strokec4 ; \cf8 \strokec8 // Total column\cf4 \cb1 \strokec4 \
\cb3       \}\cb1 \
\cb3     \});\cb1 \
\cb3     \cf5 \strokec5 htmlTable\cf4 \strokec4  += \cf7 \strokec7 '</tr>'\cf4 \strokec4 ;\cb1 \
\cb3   \});\cb1 \
\
\cb3   \cf2 \strokec2 return\cf4 \strokec4  \cf5 \strokec5 htmlTable\cf4 \strokec4 ;\cb1 \
\cb3 \}\cb1 \
\
}