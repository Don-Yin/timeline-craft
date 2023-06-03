from pptx.oxml.xmlchemy import OxmlElement


def send_backwards(slide, element):
    slide.shapes._spTree.remove(element._element)
    slide.shapes._spTree.insert(2, element._element)


def SubElement(parent, tagname, **kwargs):
    """
    Create a new element and append it to the parent element.
    """
    element = OxmlElement(tagname)
    element.attrib.update(kwargs)
    parent.append(element)
    return element


def set_shape_transparency(shape, alpha):
    """Set the transparency (alpha) of a shape"""
    ts = shape.fill._xPr.solidFill
    sF = ts.get_or_change_to_srgbClr()
    sE = SubElement(sF, "a:alpha", val=str(alpha))
