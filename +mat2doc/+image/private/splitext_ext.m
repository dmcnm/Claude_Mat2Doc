function e = splitext_ext(filename)
% SPLITEXT_EXT The extension chars of os.path.splitext(filename)[1][1:].
%
%   e = SPLITEXT_EXT(filename) returns the filename extension WITHOUT the
%   leading period and WITHOUT lowercasing, replicating CPython
%   os.path.splitext(filename)[1][1:] exactly (used by Image.ext, image.py:70).
%
%   CPython rule (posixpath/ntpath genericpath._splitext): split the BASENAME
%   at the LAST dot, but a leading run of dots is NOT a split point -- an
%   extension requires at least one non-dot character before the last dot.
%   MATLAB's fileparts does NOT match this on dotfiles (it treats ".bashrc" as
%   an extension), which is why this helper exists (Gate-2 F-1).
%
%   Vectors (Python -> this helper):
%     ".bashrc" -> ""   ".png" -> ""   "..png" -> ""   ".myimage" -> ""
%     "a.png" -> "png"  "IMG.PNG" -> "PNG" (case kept)  "a..png" -> "png"
%     ".tar.gz" -> "gz" "file." -> ""    "image" -> ""
%     "/x/y.z/a.png" -> "png" (only the basename is split; a dot in a
%     directory name does not count)
%
%   Mat2Doc infrastructure (package-private helper for +image); replicates
%   CPython os.path.splitext. No python-docx counterpart of its own.

s = char(filename);                              % accept string or char (None -> '')
% basename start: the char after the last path separator (ntpath treats both
% "/" and "\" as separators). sepPos is 0 when there is no separator.
sepPos = find(s == '/' | s == '\', 1, 'last');   % 1-based index, [] if none
if isempty(sepPos)
    sepPos = 0;
end
dotPos = find(s == '.', 1, 'last');              % last dot, 1-based, [] if none

e = "";
if ~isempty(dotPos) && dotPos > sepPos
    % Require at least one non-dot char between the basename start and the dot
    % (leading dots are not a split point).
    fi = sepPos + 1;
    while fi < dotPos
        if s(fi) ~= '.'
            e = string(s(dotPos + 1 : end));     % [1:] drops the '.' -> chars after it
            return
        end
        fi = fi + 1;
    end
end
end
