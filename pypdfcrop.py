#!/usr/bin/python
#
# Copyright 2010 Steven Robertson <steven@strobe.cc>
#
# This program is free software; you can redistribute it and/or modify
# it under the terms of the GNU General Public License version 2,
# or any later version, as published by the Free Software Foundation.

"""
A small program to [auto-]crop PDF files, similar in purpose to the 'pdfcrop'
utility. This tool appends updated page layouts to the PDF file instead of
re-rendering it entirely, which preserves more document metadata.

Requires GhostScript, unless bounding boxes are specified manually.
"""

import re
import sys
import subprocess
import shutil
from optparse import OptionParser, SUPPRESS_HELP
from itertools import *

from pyPdf import pdf
from pyPdf.generic import *

id = lambda thing: thing

# The PdfFileReader object has no external way to identify the object ID of a
# page. This hack avoids copy-pasted code at the expense of elegance.
class PdfFileReader_(pdf.PdfFileReader):
    def getObject(self, indirectReference):
        ret = super(PdfFileReader_, self).getObject(indirectReference)
        if isinstance(ret, DictionaryObject):
            ret[NameObject('/SelfHack')] = indirectReference
        return ret

def findLastXrefStart(stream):
    """Find the byte index of the last 'xref' table in the PDF."""
    stream.seek(-1024, 2)
    lines = stream.read(1024).split()
    if not lines[-3].startswith("startxref"):
        raise ValueError("Invalid PDF, I think")
    return int(lines[-2])

def foldConsecutiveList(lst, ref):
    # Add an indirect object reference to a list of lists of consecutive
    # refs, such that the list remains a list of lists of consecutive refs
    if not lst or lst[-1][-1].idnum != ref.idnum - 1:
        lst.append([ref])
    else:
        lst[-1].append(ref)
    return lst

def crop(opts, infilename):
    infile = open(infilename, 'rb')
    pdfrdr = PdfFileReader_(infile)

    if opts.bbox is None:
        if opts.bboxes is None:
            print "Determining bounding boxes for %s...\n" % infilename
            subp = subprocess.Popen(['gs', '-dBATCH', '-dNOPAUSE',
                '-sDEVICE=bbox', '-r' + opts.resolution, infilename],
                stderr=subprocess.PIPE)
            stdout, stderr = subp.communicate()
            if subp.returncode:
                raise EnvironmentError("'gs' call failed.")
        else:
            with open(opts.bboxes) as bbfile:
                stderr = bbfile.read()
        bboxes = [map(int, s.split()[1:]) for s in stderr.split('\n')
                  if s.startswith(r'%%BoundingBox:')]
    else:
        bboxes = repeat(opts.bbox)

    pages = {}
    for idx, (page, bbox) in enumerate(zip(pdfrdr.pages, bboxes)):
        pages[page.raw_get('/SelfHack')] = page
        #if page.get('/Rotate') == 90 or page.get('/Rotate') == 270:
        #    # TODO: I'm not sure this works right for 180 and 270.
        #    rbbox = [bbox[1], bbox[0], bbox[3], bbox[2]]
        #else:
        #    rbbox = bbox
        page[NameObject('/CropBox')] = ArrayObject([
                NumberObject(bbox[0]-opts.margins[0]),
                NumberObject(bbox[1]-opts.margins[1]),
                NumberObject(bbox[2]+opts.margins[2]),
                NumberObject(bbox[3]+opts.margins[3])])
        del page['/SelfHack']

    infile.seek(0)
    prev_xref_start = findLastXrefStart(infile)
    trailer = DictionaryObject(pdfrdr.trailer)
    trailer[NameObject('/Prev')] = (
            NumberObject(prev_xref_start))

    if opts.outfile is not None:
        shutil.copyfile(infilename, opts.outfile)
        outfilename = opts.outfile
    else:
        outfilename = infilename

    with open(outfilename, 'ab') as outfile:
        xref = {}
        for ref in sorted(pages.keys(), key = lambda ref: ref.idnum):
            xref[ref] = outfile.tell()
            outfile.write('%d %d obj\n' % (ref.idnum, ref.generation))
            pages[ref].writeToStream(outfile, encryption_key = None)
            outfile.write('\nendobj\n')

        current_xref_start = outfile.tell()
        outfile.write('xref \n')

        refs = sorted(xref.keys(), key = lambda ref: ref.idnum)
        subsections = reduce(foldConsecutiveList, refs, [])

        NEWLINE=' \n'
        if len(NEWLINE) == 3:
            # Windows? Anyway, just use '\n' for the xref entries
            NEWLINE='\n'

        for subsection in subsections:
            outfile.write('%d %d \n' % (subsection[0].idnum, len(subsection)))
            for ref in subsection:
                outfile.write(
                    '%010d %05d n %s' % (xref[ref], ref.generation, NEWLINE))

        outfile.write('trailer\n')
        trailer.writeToStream(outfile, encryption_key = None)
        outfile.write('\nstartxref\n%d\n%%%%EOF\n' % current_xref_start)

def main(opts, args):
    for infilename in args:
        crop(opts, infilename)

if __name__ == "__main__":
    usage=("pypdfcrop: Crop PDFs by appending instead of rewriting.\n"
           "Usage:     %prog [options] input.pdf [input2.pdf ...]\n")
    parser = OptionParser(usage=usage)
    parser.add_option("-r", "--resolution", dest="resolution", default='100',
                      help="Adjust GhostScript bounding box resolution")
    parser.add_option("-b", "--bbox", dest="bbox", default=None,
                      help="Manually set bounding box for all pages",
                      metavar='"<x1> <y1> <x2> <y2>"')
    parser.add_option('-B', "--bbox-file", dest="bboxes", default=None,
                      help=SUPPRESS_HELP, metavar="FILE")
    parser.add_option("-o", "--outfile", dest="outfile", default=None,
                      help="Set output file (default is to append to input)",
                      metavar="FILE")
    parser.add_option("-m", "--margin", dest="margins", default="0",
                      help="Pad bounding box with extra margins",
                      metavar='"<x1> [<y1> [<x2> [<y2>]]]"')
    opts, args = parser.parse_args()

    # Intelligently fill in missing margins
    opts.margins = map(int, opts.margins.split())
    if len(opts.margins) == 1:
        opts.margins.append(opts.margins[0])
    if len(opts.margins) == 2:
        opts.margins.append(opts.margins[0])
    if len(opts.margins) == 3:
        opts.margins.append(opts.margins[1])

    if opts.bbox:
        opts.bbox = map(int, opts.bbox.split())

    if not args:
        parser.print_help()
    elif len(args) > 1 and opts.outfile:
        print "Error: Cannot use --outfile with multiple input files.\n\n"
        parser.print_help()
    else:
        main(opts, args)

