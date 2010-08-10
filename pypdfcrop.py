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
import traceback
from optparse import OptionParser, SUPPRESS_HELP
from itertools import *

from pyPdf import pdf
from pyPdf.generic import *

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

def crop(opts, infilename):
    infile = open(infilename, 'rb')
    pdfrdr = PdfFileReader_(infile)

    # Ensure we can read this before continuing
    try:
        pdfrdr.getDocumentInfo()
    except Exception: # No PDF-specific exception type?
        # Some PDFs are encrypted with a blank password, try that first
        try:
            pdfrdr.decrypt('')
            pdfrdr.getDocumentInfo()
        except Exception:
            if opts.password:
                try:
                    pdfrdr.decrypt(opts.password)
                    pdfrdr.getDocumentInfo()
                except Exception:
                    traceback.print_exc()
                    print ("Got an error when decrypting PDF using given "
                           "password. Perhaps that's not it?")
            else:
                traceback.print_exc()
                print "Got an error when opening PDF. Try --password."

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
        if not bboxes:
            print "No bounding boxes detected. Dumping GS output.\n"
            print stderr
            raise EnvironmentError("'gs' call failed.")
    else:
        bboxes = repeat(opts.bbox)

    print list(pdfrdr.pages)
    pages = {}
    for idx, (page, bbox) in enumerate(zip(pdfrdr.pages, bboxes)):
        pages[page.raw_get('/SelfHack')] = page
        if opts.verbose:
            print 'Page %d: media box %s, bounding box %s' % (
                    idx, page['/MediaBox'], bbox)
        margins = (idx % 2 and opts.margins) or opts.altmargins
        page[NameObject('/CropBox')] = ArrayObject([
                NumberObject(bbox[0]-margins[0]),
                NumberObject(bbox[1]-margins[1]),
                NumberObject(bbox[2]+margins[2]),
                NumberObject(bbox[3]+margins[3])])
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

        subsections = []
        for ref in sorted(xref.keys(), key = lambda ref: ref.idnum):
            if not subsections or subsections[-1][-1].idnum != ref.idnum:
                subsections.append([ref])
            else:
                subsections[-1].append(ref)

        for subsection in subsections:
            outfile.write('%d %d \n' % (subsection[0].idnum, len(subsection)))
            for ref in subsection:
                outfile.write(('%010d %05d n \n' %
                               (xref[ref], ref.generation))[:20])

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
    parser.add_option("-v", "--verbose", dest="verbose", default=False,
                      help="Be verbose.", action="store_true")
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
    parser.add_option("-M", "--even-margin", dest="altmargins", default=None,
                      help="Specify alternate margins for even-numbered pages",
                      metavar='"<x1> [<y1> [<x2> [<y2>]]]"')
    parser.add_option("-P", "--password", dest="password", default=None,
                      help="Password to decrypt PDF.",
                      metavar="PASSWORD")
    opts, args = parser.parse_args()

    def expand_margins(margins):
        margins = map(int, margins.split())
        if len(margins) == 1:
            margins.append(margins[0])
        if len(margins) == 2:
            margins.append(margins[0])
        if len(margins) == 3:
            margins.append(margins[1])
        return margins

    opts.margins = expand_margins(opts.margins)
    if opts.altmargins is not None:
        opts.altmargins = expand_margins(opts.altmargins)
    else:
        opts.altmargins = opts.margins

    if opts.bbox:
        opts.bbox = map(int, opts.bbox.split())

    if not args:
        parser.print_help()
    elif len(args) > 1 and opts.outfile:
        print "Error: Cannot use --outfile with multiple input files.\n\n"
        parser.print_help()
    else:
        main(opts, args)

