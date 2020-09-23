VERSION 5.00
Begin VB.Form frmSymbol 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Insert Symbol"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5655
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   5655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cboFonts 
      Height          =   315
      Left            =   120
      TabIndex        =   259
      Text            =   "Combo1"
      Top             =   120
      Width           =   2655
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      Height          =   375
      Left            =   3000
      TabIndex        =   258
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "&Done"
      Height          =   375
      Left            =   4320
      TabIndex        =   257
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton cmdInsert 
      Caption         =   "&Insert"
      Height          =   375
      Left            =   1680
      TabIndex        =   256
      Top             =   4080
      Width           =   1215
   End
   Begin VB.TextBox txtCopy 
      Height          =   375
      Left            =   120
      TabIndex        =   255
      Top             =   4080
      Width           =   1455
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   255
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   254
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   254
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   253
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   253
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   252
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   252
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   251
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   251
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   250
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   250
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   249
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   249
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   248
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   248
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   247
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   247
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   246
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   246
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   245
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   245
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   244
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   244
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   243
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   243
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   242
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   242
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   241
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   241
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   240
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   240
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   239
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   239
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   238
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   238
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   237
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   237
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   236
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   236
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   235
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   235
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   234
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   234
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   233
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   233
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   232
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   232
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   231
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   231
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   230
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   230
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   229
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   229
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   228
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   228
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   227
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   227
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   226
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   226
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   225
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   225
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   224
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   224
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   223
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   223
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   222
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   222
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   221
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   221
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   220
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   220
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   219
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   219
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   218
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   218
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   217
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   217
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   216
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   216
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   215
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   215
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   214
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   214
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   213
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   213
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   212
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   212
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   211
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   211
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   210
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   210
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   209
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   209
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   208
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   208
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   207
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   207
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   206
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   206
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   205
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   205
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   204
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   204
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   203
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   203
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   202
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   202
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   201
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   201
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   200
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   200
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   199
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   199
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   198
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   198
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   197
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   197
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   196
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   196
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   195
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   195
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   194
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   194
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   193
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   193
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   192
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   192
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   191
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   191
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   190
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   190
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   189
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   189
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   188
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   188
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   187
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   187
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   186
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   186
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   185
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   185
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   184
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   184
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   183
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   183
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   182
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   182
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   181
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   181
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   180
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   180
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   179
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   179
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   178
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   178
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   177
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   177
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   176
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   176
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   175
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   175
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   174
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   174
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   173
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   173
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   172
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   172
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   171
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   171
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   170
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   170
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   169
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   169
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   168
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   168
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   167
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   167
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   166
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   166
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   165
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   165
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   164
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   164
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   163
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   163
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   162
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   162
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   161
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   161
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   160
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   160
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   159
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   159
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   158
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   158
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   157
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   157
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   156
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   156
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   155
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   155
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   154
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   154
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   153
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   153
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   152
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   152
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   151
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   151
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   150
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   150
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   149
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   149
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   148
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   148
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   147
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   147
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   146
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   146
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   145
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   145
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   144
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   144
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   143
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   143
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   142
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   142
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   141
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   141
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   140
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   140
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   139
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   139
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   138
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   138
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   137
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   137
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   136
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   136
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   135
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   135
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   134
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   134
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   133
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   133
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   132
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   132
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   131
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   131
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   130
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   130
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   129
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   129
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   128
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   128
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   127
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   127
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   126
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   126
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   125
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   125
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   124
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   124
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   123
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   123
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   122
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   122
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   121
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   121
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   120
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   120
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   119
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   119
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   118
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   118
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   117
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   117
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   116
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   116
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   115
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   115
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   114
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   114
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   113
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   113
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   112
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   112
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   111
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   111
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   110
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   110
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   109
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   109
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   108
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   108
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   107
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   107
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   106
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   106
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   105
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   105
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   104
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   104
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   103
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   103
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   102
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   102
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   101
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   101
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   100
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   100
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   99
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   99
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   98
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   98
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   97
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   97
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   96
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   96
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   95
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   95
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   94
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   94
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   93
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   93
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   92
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   92
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   91
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   91
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   90
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   90
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   89
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   89
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   88
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   88
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   87
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   87
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   86
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   86
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   85
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   85
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   84
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   84
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   83
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   83
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   82
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   82
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   81
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   81
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   80
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   80
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   79
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   79
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   78
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   78
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   77
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   77
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   76
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   76
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   75
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   75
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   74
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   74
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   73
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   73
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   72
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   72
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   71
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   71
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   70
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   70
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   69
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   69
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   68
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   68
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   67
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   67
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   66
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   66
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   65
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   65
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   64
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   64
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   63
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   63
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   62
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   62
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   61
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   61
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   60
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   60
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   59
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   59
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   58
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   58
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   57
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   57
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   56
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   56
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   55
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   55
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   54
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   54
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   53
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   53
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   52
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   52
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   51
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   51
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   50
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   50
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   49
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   48
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   47
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   46
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   45
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   44
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   43
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   42
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   41
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   40
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   39
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   38
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   37
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   36
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   35
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   34
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   33
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   32
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   31
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   30
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   29
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   28
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   27
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   26
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   25
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   24
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   23
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   22
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   21
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   20
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   19
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   18
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   17
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   16
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   15
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   14
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   13
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   12
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   11
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   10
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   9
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   8
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   7
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   6
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   5
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   4
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   3
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   2
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdSymbol 
      Height          =   255
      Index           =   1
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   600
      Width           =   255
   End
End
Attribute VB_Name = "frmSymbol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Integer

Private Sub cboFonts_Click()
On Error Resume Next
    For i = 1 To 255
        cmdSymbol(i).FontName = cboFonts.Text
    Next i
End Sub

Private Sub cboFonts_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then cboFonts_Click
End Sub

Private Sub cmdClear_Click()
txtCopy.Text = ""
End Sub

Private Sub cmdDone_Click()
Unload Me
End Sub

Private Sub cmdInsert_Click()
frmMain.ActiveForm.rtfText.SelText = txtCopy.Text
frmMain.ActiveForm.rtfText.SetFocus
End Sub

Private Sub cmdSymbol_Click(Index As Integer)
txtCopy.Text = txtCopy.Text & Chr(Index)
End Sub

Private Sub Form_Load()
Dim j As Integer
Me.Show
MakeTop Me
    For i = 1 To 255
        cmdSymbol(i).Left = (i Mod 21) * 255
        cmdSymbol(i).Top = j * 255 + 600
            If i Mod 21 = 0 Then
                cmdSymbol(i).Left = 255 * 20
                cmdSymbol(i).Top = j * 255 + 600
                j = j + 1
            End If
        cmdSymbol(i).Caption = Chr(i)
    Next i
    For i = 0 To frmMain.cboFonts.ListCount
            If frmMain.cboFonts.List(i) = "Easter Egg" Then GoTo NextI
        cboFonts.AddItem frmMain.cboFonts.List(i)
NextI:
    Next i
cboFonts.Text = frmMain.cboFonts.Text
End Sub
