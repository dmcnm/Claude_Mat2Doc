classdef ImageParts < handle
% IMAGEPARTS Collection of ImagePart objects corresponding to images in a package.
%
%   A package-level collection (one per Package, held by Package.image_parts) that
%   de-duplicates image parts by SHA1: adding the same image bytes twice reuses
%   the one ImagePart. Populated at open by Package._gather_image_parts (every
%   existing internal IMAGE relationship's target part) and grown on demand by
%   get_or_add_image_part (the add_picture path).
%
%   REFERENCE SEMANTICS (design.md section 2): a handle class -- a mutable
%   collection shared by reference, exactly like the Python object with its
%   mutable `_image_parts` list. H5 identity: membership (__contains__) and the
%   iter_rels-driven gather compare parts by handle identity.
%
%   VERIFY-COLLECTION (design.md section 2): the Python container surface
%   (__contains__ / __iter__ / __len__) is ported as EXPLICIT methods
%   (contains / to_array / len_), matching the InlineShapes getitem_/to_array/len_
%   and Rows_/Columns_ precedent. Dunder mapping: `x in image_parts` ->
%   image_parts.contains(x); `for p in image_parts` -> `for p =
%   image_parts.to_array()`; `len(image_parts)` -> image_parts.len_().
%
%   UNDERSCORE ROTATION (design.md section 2): private `_image_parts` ->
%   image_parts_; the private helpers `_add_image_part` -> add_image_part_,
%   `_get_by_sha1` -> get_by_sha1_, `_next_image_partname` -> next_image_partname_.
%
%   Ported from python-docx v1.2.0: src/docx/package.py::ImageParts (lines 50-111)

    properties (Access = private)
        image_parts_   % _image_parts (image.py: list[ImagePart]) -- an ImagePart array
    end

    methods
        function obj = ImageParts()
            % Ported from ImageParts.__init__ (package.py 53-54): empty list.
            obj.image_parts_ = mat2doc.parts.ImagePart.empty(1, 0);
        end

        function tf = contains(obj, item)
            % __contains__ (package.py 56-57): membership by handle identity
            %   (Python list `in`, which for Part uses identity `==`). Empty
            %   collection -> false.
            if isempty(obj.image_parts_)
                tf = false;
                return
            end
            tf = any(obj.image_parts_ == item);
        end

        function arr = to_array(obj)
            % __iter__ (package.py 59-60): the ImagePart array, for the
            %   `for p in image_parts` idiom. Returns the live array (a value copy
            %   of the handle vector; the ImagePart handles are shared).
            arr = obj.image_parts_;
        end

        function n = len_(obj)
            % __len__ (package.py 62-63).
            n = numel(obj.image_parts_);
        end

        function append(obj, item)
            % append (package.py 65-66): add `item` to the collection.
            obj.image_parts_(end + 1) = item;
        end

        function image_part = get_or_add_image_part(obj, image_descriptor)
            % get_or_add_image_part (package.py 68-78): the ImagePart for
            %   `image_descriptor`, created if a SHA1 match is not already present.
            %   Python:
            %     image = Image.from_file(image_descriptor)
            %     matching_image_part = self._get_by_sha1(image.sha1)
            %     if matching_image_part is not None:
            %         return matching_image_part
            %     return self._add_image_part(image)
            image = mat2doc.image.Image.from_file(image_descriptor);
            matching_image_part = obj.get_by_sha1_(image.sha1);
            if ~isequal(matching_image_part, [])   % Python: if ... is not None
                image_part = matching_image_part;
                return
            end
            image_part = obj.add_image_part_(image);
        end
    end

    methods (Access = private)
        function image_part = add_image_part_(obj, image)
            % _add_image_part (package.py 80-85): a new ImagePart from `image`,
            %   appended to the collection. Python:
            %     partname = self._next_image_partname(image.ext)
            %     image_part = ImagePart.from_image(image, partname)
            %     self.append(image_part)
            %     return image_part
            partname = obj.next_image_partname_(image.ext);
            image_part = mat2doc.parts.ImagePart.from_image(image, partname);
            obj.append(image_part);
        end

        function image_part = get_by_sha1_(obj, sha1)
            % _get_by_sha1 (package.py 87-93): the image part whose SHA1 matches
            %   `sha1`, or [] (None). Python:
            %     for image_part in self._image_parts:
            %         if image_part.sha1 == sha1: return image_part
            %     return None
            parts = obj.image_parts_;
            for i = 1:numel(parts)
                if parts(i).sha1 == sha1   % Python: image_part.sha1 == sha1
                    image_part = parts(i);
                    return
                end
            end
            image_part = [];   % Python: return None (H3)
        end

        function uri = next_image_partname_(obj, ext)
            % _next_image_partname (package.py 95-110): the next available image
            %   partname "/word/media/image<n>.<ext>", reusing unused numbers. The
            %   number is unique WITHOUT regard to extension. Python:
            %     used_numbers = [image_part.partname.idx for image_part in self]
            %     for n in range(1, len(self) + 1):
            %         if n not in used_numbers: return image_partname(n)
            %     return image_partname(len(self) + 1)
            %   image_partname(n) = PackURI("/word/media/image%d.%s" % (n, ext)).
            %   H1: `n` is a partname NUMBER (data), not a 0/1 index -- ported
            %   verbatim (image parts always carry a numeric partname.idx).
            parts = obj.image_parts_;
            used_numbers = double.empty(1, 0);
            for i = 1:numel(parts)
                % PackURI.idx is [] (None) for a NON-numbered media partname
                % (packuri.py 61-75), e.g. "/word/media/logo.png". Python
                % (package.py:106) appends None to used_numbers harmlessly -- a
                % None entry never equals a candidate integer, so `n not in
                % used_numbers` skips it. Appending [] to a MATLAB numeric array
                % would delete an element (singleSubscript error), so skip the
                % None entries: a list of just the real integers gives the
                % IDENTICAL next-partname result (Gate-2 DEFECT-1 fix).
                idx = parts(i).partname().idx;
                if ~isequal(idx, [])   % Python: None append is a no-op for the test
                    used_numbers(end + 1) = idx; %#ok<AGROW>
                end
            end
            n_self = numel(parts);                 % len(self)
            for n = 1:n_self                        % Python: range(1, len(self)+1)
                if ~any(used_numbers == n)          % Python: if n not in used_numbers
                    uri = image_partname_(n, ext);
                    return
                end
            end
            uri = image_partname_(n_self + 1, ext); % Python: image_partname(len(self)+1)
        end
    end
end

% ------------------------------------------------------------------------
% File-local: the nested image_partname(n) closure (package.py 103-104).
% ------------------------------------------------------------------------
function uri = image_partname_(n, ext)
% Python: PackURI("/word/media/image%d.%s" % (n, ext))
uri = mat2doc.opc.PackURI("/word/media/image" + string(n) + "." + ext);
end
