# -*- coding: utf-8 -*-
# pyRevit skript: Report vybraných prvků + info o souboru, verzi, čas a uživateli.
# Do schránky vloží kombinaci textového reportu a screenshotu aktivního okna jako HTML (pro Teams, Outlook apod.)

import clr, ctypes, base64, os
clr.AddReference('System')
clr.AddReference('System.Drawing')
clr.AddReference('System.Windows.Forms')

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import TaskDialog
from System import DateTime
from System.Drawing import Bitmap, Graphics, Imaging
from System.Windows.Forms import Clipboard, DataObject, DataFormats, Form, TextBox, Button, Label, DialogResult, FormStartPosition
from System.IO import MemoryStream

uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document

# --- Win32 imports ---
user32 = ctypes.windll.user32
gdi32 = ctypes.windll.gdi32

# -------- helpers --------
def _png_bytes_from_bitmap(bmp):
    ms = MemoryStream()
    bmp.Save(ms, Imaging.ImageFormat.Png)
    arr = ms.ToArray()
    ms.Dispose()
    return arr

def html_escape(text):
    replacements = {"&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;"}
    escaped = "".join(replacements.get(c, c) for c in text)
    def to_entity(c):
        return "&#{};".format(ord(c)) if ord(c) > 127 else c
    return "".join(to_entity(c) for c in escaped)

# obrázek až ZA textem
def _build_html_with_image_and_caption(png_bytes, caption_text):
    b64 = base64.b64encode(bytearray(png_bytes)).decode("ascii")
    data_uri = "data:image/png;base64," + b64
    safe_caption = html_escape(caption_text)
    html_fragment = (
        '<div>'
        '<meta charset="utf-8">'
        '<pre style="font-family:Segoe UI, Consolas, monospace; font-size:10pt; white-space:pre-wrap;">{0}</pre>'
        '<br/>'
        '<img src="{1}" alt="screenshot" />'
        '</div>'
    ).format(safe_caption, data_uri)
    prefix = (
        "Version:0.9\r\n"
        "StartHTML:{st_html:08d}\r\n"
        "EndHTML:{end_html:08d}\r\n"
        "StartFragment:{st_frag:08d}\r\n"
        "EndFragment:{end_frag:08d}\r\n"
    )
    frag_start = "<html><head><meta charset='utf-8'></head><body><!--StartFragment-->"
    frag_end = "<!--EndFragment--></body></html>"
    stub = prefix.format(st_html=0, end_html=0, st_frag=0, end_frag=0)
    st_html = len(stub)
    st_frag = st_html + len(frag_start)
    html = frag_start + html_fragment + frag_end
    end_frag = st_frag + len(html_fragment)
    end_html = st_html + len(html)
    header = prefix.format(st_html=st_html, end_html=end_html, st_frag=st_frag, end_frag=end_frag)
    return header + html

def _put_on_clipboard(img_bmp, caption_text):
    html_payload = _build_html_with_image_and_caption(_png_bytes_from_bitmap(img_bmp), caption_text)
    obj = DataObject()
    obj.SetData(DataFormats.Html, html_payload)
    obj.SetImage(img_bmp)
    Clipboard.SetDataObject(obj, True)

# -------- active window capture --------
class RECT(ctypes.Structure):
    _fields_ = [("left", ctypes.c_long), ("top", ctypes.c_long), ("right", ctypes.c_long), ("bottom", ctypes.c_long)]

def grab_active_window_bitmap():
    hwnd = user32.GetForegroundWindow()
    rect = RECT()
    user32.GetWindowRect(hwnd, ctypes.byref(rect))
    w, h = rect.right - rect.left, rect.bottom - rect.top
    bmp = Bitmap(w, h)
    g = Graphics.FromImage(bmp)
    hdc_dest = g.GetHdc()
    hdc_src = user32.GetWindowDC(hwnd)
    gdi32.BitBlt(hdc_dest, 0, 0, w, h, hdc_src, 0, 0, 0x00CC0020)
    g.ReleaseHdc(hdc_dest)
    g.Dispose()
    user32.ReleaseDC(hwnd, hdc_src)
    return bmp

# -------- vlastní InputBox pro popis problému --------
def ask_user_problem(prompt="Popište prosím svůj problém:", title="Popis problému"):
    form = Form()
    form.Text = title
    form.Width = 800   # zdvojnásobeno
    form.Height = 400  # zdvojnásobeno
    form.StartPosition = FormStartPosition.CenterScreen

    label = Label()
    label.Text = prompt
    label.Top = 10
    label.Left = 10
    label.Width = 760
    form.Controls.Add(label)

    textbox = TextBox()
    textbox.Multiline = True
    textbox.Width = 760   # zdvojnásobeno
    textbox.Height = 280  # zdvojnásobeno
    textbox.Top = 30
    textbox.Left = 10
    form.Controls.Add(textbox)

    button_ok = Button()
    button_ok.Text = "OK"
    button_ok.Top = 320
    button_ok.Left = 680
    button_ok.DialogResult = DialogResult.OK
    form.Controls.Add(button_ok)

    form.AcceptButton = button_ok


    result = form.ShowDialog()
    if result == DialogResult.OK:
        return textbox.Text
    return ""

# -------- main report logic --------
selection_ids = uidoc.Selection.GetElementIds()
elements_info = []

for el_id in selection_ids:
    el = doc.GetElement(el_id)
    if el:
        cat_name = el.Category.Name if el.Category else "Bez kategorie"
        try:
            type_id = el.GetTypeId()
            el_type = doc.GetElement(type_id)
            family_name = el_type.FamilyName
        except:
            family_name = "No FamilyName atribute"
        type_name = el.Name
        # Opravený formát: ID – Kategorie – Rodina – Typ
        elements_info.append("{0} – {1} – {2} – {3}".format(el.Id, cat_name, family_name, type_name))

ids_text = "\n".join(elements_info) if elements_info else "Žádné prvky nebyly vybrány."

# --- soubor a cesta ---
central_bool = None
full_path = None
if doc.IsWorkshared:
    try:
        central_path_obj = doc.GetWorksharingCentralModelPath()
        if central_path_obj is not None:
            full_path = ModelPathUtils.ConvertModelPathToUserVisiblePath(central_path_obj)
            central_bool = True
    except:
        full_path = None
else:
    try:
        local_path = doc.PathName
        if local_path is not None:
            full_path = local_path
            central_bool = False
    except:
        full_path = None

is_central = "Ano" if central_bool else "Ne"

if not full_path or full_path.strip() == "":
    file_name = "Soubor není uložen."
    file_path = "Cesta neexistuje."
else:
    if full_path.startswith(r"\\10.78.0.100\projekty"):
        full_path = full_path.replace(r"\\10.78.0.100\projekty", "x:")
    file_name = os.path.basename(full_path)
    file_path = os.path.dirname(full_path)
    if not file_path.endswith("\\"):
        file_path += "\\"

# --- aktuální uživatel ---
user_name = getattr(doc.Application, "Username", "Neznámý uživatel")
revit_version = doc.Application.SubVersionNumber
current_time = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")

# --- dotaz na problém ---
user_issue = ask_user_problem()
if user_issue.strip() == "":
    user_issue = "Uživatel nevyplnil popis."

# --- sestavení reportu ---
output_text = (
    "Report: {0}\n\n"
    "Název: {1}\n"
    "Centrální model: {2}\n"
    "Cesta: {3}\n"
    "Uživatel: {4}\n"
    "Revit: {5}\n\n"
    "Popis problému:\n{6}\n\n"
    "Označené prvky:\n{7}"
).format(current_time, file_name, is_central, file_path, user_name, revit_version, user_issue, ids_text)

# --- screenshot + clipboard ---
bmp = grab_active_window_bitmap()
_put_on_clipboard(bmp, output_text)
bmp.Dispose()

# --- TaskDialog ---
TaskDialog.Show("Hotovo", "Report je připraven pod Ctrl + V. Pošlete ho přes Teamsy svému BIM koordinátorovi.")
