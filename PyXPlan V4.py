#!/usr/bin/env python3
#
#  pyxplan V4.py
#  
#  Copyright 2024 olivier <olivier@olivier-MS-7817>
#  
#  This program is free software; you can redistribute it and/or modify
#  it under the terms of the GNU General Public License as published by
#  the Free Software Foundation; either version 2 of the License, or
#  (at your option) any later version.
#  
#  This program is distributed in the hope that it will be useful,
#  but WITHOUT ANY WARRANTY; without even the implied warranty of
#  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#  GNU General Public License for more details.
#  
#  You should have received a copy of the GNU General Public License
#  along with this program; if not, write to the Free Software
#  Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston,
#  MA 02110-1301, USA.
#  
#  

import os
import sys
import tkgui_pyxplan as gpxp

def main(args):
  iconfile=os.path.join(os.getcwd(),"Biplan.png")
  app=gpxp.MainWindow('PyXPlan V4', (600,800), iconfile)
  return 0


if __name__ == '__main__':
  sys.exit(main(sys.argv[1:]))
