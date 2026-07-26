classdef RELATIONSHIP_TYPE
% RELATIONSHIP_TYPE Open XML relationship-type URIs.
%
%   Access as mat2doc.opc.RELATIONSHIP_TYPE.OFFICE_DOCUMENT etc. Mirrors the
%   Python class attributes (class RELATIONSHIP_TYPE with class-level
%   constants). Each value is transcribed verbatim from the docx source; the
%   Python implicit-string-concatenation entries (values split across source
%   lines for line length) are joined into the single string they denote --
%   NOTE PIVOT_CACHE_RECORDS interleaves ".../relationships" + "/spreadsheetml/
%   pivotCacheRecords" (an infix that differs from the sibling entries), joined
%   exactly here.
%
%   Example:
%       disp(mat2doc.opc.RELATIONSHIP_TYPE.OFFICE_DOCUMENT)
%       % "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
%
%   Ported from python-docx v1.2.0: src/docx/opc/constants.py::RELATIONSHIP_TYPE
%   (lines 178-306)

    properties (Constant)
        AUDIO = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio"
        A_F_CHUNK = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk"
        CALC_CHAIN = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain"
        CERTIFICATE = "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/certificate"
        CHART = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"
        CHARTSHEET = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet"
        CHART_USER_SHAPES = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartUserShapes"
        COMMENTS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
        COMMENT_AUTHORS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/commentAuthors"
        CONNECTIONS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/connections"
        CONTROL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/control"
        CORE_PROPERTIES = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"
        CUSTOM_PROPERTIES = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties"
        CUSTOM_PROPERTY = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customProperty"
        CUSTOM_XML = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml"
        CUSTOM_XML_PROPS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps"
        DIAGRAM_COLORS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramColors"
        DIAGRAM_DATA = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData"
        DIAGRAM_LAYOUT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramLayout"
        DIAGRAM_QUICK_STYLE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramQuickStyle"
        DIALOGSHEET = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/dialogsheet"
        DRAWING = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"
        ENDNOTES = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes"
        EXTENDED_PROPERTIES = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties"
        EXTERNAL_LINK = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink"
        FONT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/font"
        FONT_TABLE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable"
        FOOTER = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"
        FOOTNOTES = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes"
        GLOSSARY_DOCUMENT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/glossaryDocument"
        HANDOUT_MASTER = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/handoutMaster"
        HEADER = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
        HYPERLINK = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
        IMAGE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
        NOTES_MASTER = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster"
        NOTES_SLIDE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide"
        NUMBERING = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering"
        OFFICE_DOCUMENT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
        OLE_OBJECT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject"
        ORIGIN = "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin"
        PACKAGE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package"
        PIVOT_CACHE_DEFINITION = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition"
        PIVOT_CACHE_RECORDS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/spreadsheetml/pivotCacheRecords"
        PIVOT_TABLE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable"
        PRES_PROPS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps"
        PRINTER_SETTINGS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings"
        QUERY_TABLE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/queryTable"
        REVISION_HEADERS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/revisionHeaders"
        REVISION_LOG = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/revisionLog"
        SETTINGS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"
        SHARED_STRINGS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"
        SHEET_METADATA = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata"
        SIGNATURE = "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature"
        SLIDE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"
        SLIDE_LAYOUT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"
        SLIDE_MASTER = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster"
        SLIDE_UPDATE_INFO = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideUpdateInfo"
        STYLES = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
        TABLE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table"
        TABLE_SINGLE_CELLS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableSingleCells"
        TABLE_STYLES = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles"
        TAGS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/tags"
        THEME = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
        THEME_OVERRIDE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/themeOverride"
        THUMBNAIL = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail"
        USERNAMES = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/usernames"
        VIDEO = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/video"
        VIEW_PROPS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps"
        VML_DRAWING = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing"
        VOLATILE_DEPENDENCIES = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/volatileDependencies"
        WEB_SETTINGS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings"
        WORKSHEET_SOURCE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheetSource"
        XML_MAPS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/xmlMaps"
    end
end
