# $Id$
#
# Author: LWW
# Date: 3 May 2023
# 
# Plugin to create GCode from Gerber and Excellon files using the PCB2GCode utility.

import os
import glob
import re
import json
import subprocess

import Utils
from ToolsPage import Plugin
import tkinter as tk
from tkinter import ttk, filedialog

__author__ = "LWW"
__email__ = "lloyd.wwynn@gmail.com"

__name__ = _("PCB2Gcode")
__version__ = "0.0.1"

any_filetypes = ("Any","*.*")
gerber_filetypes = ("Gerber", "*.gbr")
drill_filetypes = ("Excellon", "*.drl")

PROJECT_SETTINGS_FILE = "PCB2GCode.json"

DRILL_OUTPUT_FILE = "drill.ngc"
DRILL_CLONE_FILE = "drill_{}.ngc"

# Project Settings Tags
FRONT_COPPER = "FRONT_COPPER"
BACK_COPPER = "BACK_COPPER"
FRONT_ENGRAVING = "FRONT_ENGRAVING"
BACK_ENGRAVING = "BACK_ENGRAVING"
OUTLINE = "OUTLINE"
DRILLING = "DRILLING"
REMOVE_TOOL_CHANGES = "REMOVE_TOOL_CHANGES"

def files_exist(files):
    """Check the list of files supplied to ensure that they all exist

    Args:
        files[string]: Each entry is a full path to a file.

    Returns:
        bool: `True` if all files exist, `False` otherwise.
    """
    for f in files:
        if not os.path.isfile(f):
            return False
    return True

def remove_files(path, file_pattern):
    """Remove all files from folder specified in path matching the file name pattern.

    Args:
        path (string): Path-like string
        
        file_pattern (string): File name pattern, eg *.txt
    """
    file_list = glob.glob(os.path.join(path, file_pattern))
    for f in file_list:
        try:
            os.remove(f)
        except:
            print("Error while deleting file :", f)


# =============================================================================
# PCB2GCode class
# =============================================================================
class PCB2GCode:
    def __init__(self, plugin, app, name="PCB2GCode"):
        self.name = name
        self.plugin = plugin # the plugin context provides access to the plugin variables 
        self.app = app # the app context provides all application data and the main UI parent window
        
        self.project_path = self.plugin["ProjectPath"]
        self.project_folder = tk.StringVar(value=self.project_path)
    
        # define data that will be persisted for each project    
        self.project_settings = {
            FRONT_COPPER: tk.StringVar(),
            BACK_COPPER: tk.StringVar(),
            FRONT_ENGRAVING: tk.StringVar(),
            BACK_ENGRAVING: tk.StringVar(),
            OUTLINE: tk.StringVar(),
            DRILLING: tk.StringVar(),
            REMOVE_TOOL_CHANGES: tk.StringVar(value="")           
        }
        self.load_project_settings()
        
        self.ui = self.create_ui()
            
    def create_ui(self):
        """ Build the dialog window for the PCB2GCode user interface.
        """
        file_icon = Utils.icons["load"]
        ui = tk.Toplevel(self.app)
        ui.title("PCB2GCode GCode Generator")
        ui.protocol("WM_DELETE_WINDOW", self.dismiss) # intercept the close button
        ui.geometry("+{}+{}".format(self.app.winfo_x()+100, self.app.winfo_y()+100))
        
        ui.columnconfigure(0, weight=1)
        ui.rowconfigure(0, weight=0)
        ui.rowconfigure(1, weight=0)
        
        # Create body frame for path data entry widgets
        self.body_frame = ttk.Frame(ui)
        self.body_frame.grid(column=0, row=0, sticky=tk.EW)
        self.body_frame.columnconfigure(0, weight=0)
        self.body_frame.columnconfigure(1, weight=1)
        self.body_frame.columnconfigure(2, weight=0)
        
        # Create frame for options widgets
        self.options_frame = ttk.Frame(ui)
        self.options_frame.grid(column=0, row=1, sticky=tk.EW)
        self.options_frame.columnconfigure(0, weight=1)
        self.options_frame.columnconfigure(1, weight=1)
        
        # Create frame for buttons
        self.button_frame = ttk.Frame(ui)
        self.button_frame.grid(column=0, row=2)
        self.button_frame.columnconfigure(0, weight=1)
        self.button_frame.columnconfigure(1, weight=1)
        
        # Create rows of data entry elements

        # Project folder entry/selection
        row = 0
        self.project_folder_label = ttk.Label(self.body_frame, text="Project Folder")
        self.project_folder_label.grid(column=0, row=row, sticky=tk.E)
        self.project_folder_entry = ttk.Entry(self.body_frame, width=60, textvariable=self.project_folder)
        self.project_folder_entry.grid(column=1, row=row, sticky=tk.EW)
        self.project_folder_btn = ttk.Button(self.body_frame, image=file_icon, command=self.ask_project_folder)
        self.project_folder_btn.grid(column=2, row=row, sticky=tk.W)
        self.project_folder.trace_add("write", self.change_project_path)
                        
        # Front copper entry
        row += 1
        self.front_copper_label = ttk.Label(self.body_frame, text="Front Copper Gerber File")
        self.front_copper_label.grid(column=0, row=row, sticky=tk.E)
        self.front_copper_entry = ttk.Entry(self.body_frame, width=60, textvariable=self.project_settings[FRONT_COPPER])
        self.front_copper_entry.grid(column=1, row=row, sticky=tk.EW)
        self.front_copper_btn = ttk.Button(self.body_frame, image=file_icon, command=self.ask_front_filename)
        self.front_copper_btn.grid(column=2, row=row, sticky=tk.W)
        
        # Back copper entry
        row += 1
        self.back_copper_label = ttk.Label(self.body_frame, text="Back Copper Gerber File")
        self.back_copper_label.grid(column=0, row=row, sticky=tk.E)
        self.back_copper_entry = ttk.Entry(self.body_frame, width=60, textvariable=self.project_settings[BACK_COPPER])
        self.back_copper_entry.grid(column=1, row=row, sticky=tk.EW)
        self.back_copper_btn = ttk.Button(self.body_frame, image=file_icon, command=self.ask_back_filename)
        self.back_copper_btn.grid(column=2, row=row, sticky=tk.W)
        
        # Front engraving entry
        row += 1
        self.front_engraving_label = ttk.Label(self.body_frame, text="Front Engraving Gerber File")
        self.front_engraving_label.grid(column=0, row=row, sticky=tk.E)
        self.front_engraving_entry = ttk.Entry(self.body_frame, width=60, textvariable=self.project_settings[FRONT_ENGRAVING])
        self.front_engraving_entry.grid(column=1, row=row, sticky=tk.EW)
        self.front_engraving_btn = ttk.Button(self.body_frame, image=file_icon, command=self.ask_front_engraving_filename)
        self.front_engraving_btn.grid(column=2, row=row, sticky=tk.W)
        
        # Back engraving entry
        row += 1
        self.back_engraving_label = ttk.Label(self.body_frame, text="Back Engraving Gerber File")
        self.back_engraving_label.grid(column=0, row=row, sticky=tk.E)
        self.back_engraving_entry = ttk.Entry(self.body_frame, width=60, textvariable=self.project_settings[BACK_ENGRAVING])
        self.back_engraving_entry.grid(column=1, row=row, sticky=tk.EW)
        self.back_engraving_btn = ttk.Button(self.body_frame, image=file_icon, command=self.ask_back_engraving_filename)
        self.back_engraving_btn.grid(column=2, row=row, sticky=tk.W)
        
        # Outline milling entry
        row += 1
        self.outline_label = ttk.Label(self.body_frame, text="Outline Gerber File")
        self.outline_label.grid(column=0, row=row, sticky=tk.E)
        self.outline_entry = ttk.Entry(self.body_frame, width=60, textvariable=self.project_settings[OUTLINE])
        self.outline_entry.grid(column=1, row=row, sticky=tk.EW)
        self.outline_btn = ttk.Button(self.body_frame, image=file_icon, command=self.ask_outline_filename)
        self.outline_btn.grid(column=2, row=row, sticky=tk.W)
        
        # Drilling entry
        row += 1
        self.drilling_label = ttk.Label(self.body_frame, text="Drilling (Excellon) File")
        self.drilling_label.grid(column=0, row=row, sticky=tk.E)
        self.drilling_entry = ttk.Entry(self.body_frame, width=60, textvariable=self.project_settings[DRILLING])
        self.drilling_entry.grid(column=1, row=row, sticky=tk.EW)
        self.drilling_btn = ttk.Button(self.body_frame, image=file_icon, command=self.ask_drilling_filename)
        self.drilling_btn.grid(column=2, row=row, sticky=tk.W)
        
        # Options
        row = 0
        self.checkbox = ttk.Checkbutton(self.options_frame, text="Remove tool changes",
                command=None,
                variable=self.project_settings[REMOVE_TOOL_CHANGES],
                onvalue="Y",
                offvalue="")
        self.checkbox.grid(column=1, row=row, padx=5, pady=5)
        
        # Buttons
        row = 0
        self.cancel_btn = ttk.Button(self.button_frame, text="Cancel", command=self.dismiss)
        self.cancel_btn.grid(column=1, row=row, padx=10, pady=5)
        ui.bind("<Key-Escape>", lambda e: self.cancel_btn.invoke())
        self.generate_btn = ttk.Button(self.button_frame, text="Generate", command=self.generate_gcode)
        self.generate_btn.grid(column=2, row=row, padx=10, pady=5)
        ui.bind("<Key-Return>", lambda e: self.generate_btn.invoke())
        
        return ui
            
    def dismiss(self):
        print("Saving project settings and closing...")
        self.save_project_settings()
        self.ui.grab_release()
        self.ui.destroy()

    def show(self):        
        print("Opening the PCB2GCode dialog")   
        self.ui.transient(self.app)   # dialog window is a child of the app instance
        self.ui.wait_visibility() # can't grab until window appears, so we wait
        self.ui.grab_set()        # ensure all input goes to this window
        self.ui.focus_set()       # set the UI focus to this window
        self.ui.wait_window()     # block until window is destroyed (i.e. modal dialog)
        
    def ask_project_folder(self):
        folder_path = filedialog.askdirectory(initialdir=self.project_folder.get())
        if folder_path:
            self.project_folder.set(folder_path)

    def ask_filename(self, item, filetypes):
        filepath = filedialog.askopenfilename(filetypes=[filetypes], 
                                              initialdir=self.project_path,
                                              initialfile=item.get())
        filename = os.path.basename(filepath) if filepath else ""
        if filename:
            item.set(filename)
                    
    def ask_front_filename(self):
        self.ask_filename(self.project_settings[FRONT_COPPER], gerber_filetypes)
            
    def ask_back_filename(self):
        self.ask_filename(self.project_settings[BACK_COPPER], gerber_filetypes)
            
    def ask_front_engraving_filename(self):
        self.ask_filename(self.project_settings[FRONT_ENGRAVING], gerber_filetypes)
            
    def ask_back_engraving_filename(self):
        self.ask_filename(self.project_settings[BACK_ENGRAVING], gerber_filetypes)
            
    def ask_outline_filename(self):
        self.ask_filename(self.project_settings[OUTLINE], gerber_filetypes)
            
    def ask_drilling_filename(self):
        self.ask_filename(self.project_settings[DRILLING], drill_filetypes)

    def change_project_path(self, *args):
        new_project_path = self.project_folder.get()
        if new_project_path and (new_project_path != self.project_path):
            self.save_project_settings()
            self.project_path = new_project_path
            self.load_project_settings()     
            
    def save_project_settings(self):
        """Save the current project settings to a JSON file in the project folder
        """
        if self.project_path and os.path.isdir(self.project_path):
            self.plugin["ProjectPath"] = self.project_path
            settings_filename = os.path.join(self.project_path, PROJECT_SETTINGS_FILE)
            # Copy the current settings to a dictionary
            settings = {key: value.get() for (key, value) in self.project_settings.items()}
            # Check if any of the settings are actually used
            settings_used = False
            for tag in settings.keys():
                if settings[tag]:
                    settings_used = True
            if settings_used:
                with open(settings_filename, "w") as fp:
                    json.dump(settings, fp, sort_keys=True, indent=4, ensure_ascii=False)
    
    def load_project_settings(self):
        """Load the saved settings from JSON to the project_settings dictionary
        """
        if self.project_path and os.path.isdir(self.project_path):
            settings_filename = os.path.join(self.project_path, PROJECT_SETTINGS_FILE)
            if os.path.isfile(settings_filename):
                try:
                    with open(settings_filename, "r") as fp:
                        saved_settings = json.load(fp)
                    for tag in self.project_settings.keys():
                        if tag in saved_settings:
                            self.project_settings[tag].set(saved_settings[tag])
                except:
                    print("Badly formed settings info - ignored")
            else:
                # no project settings found, blank out all the settings values
                for tag in self.project_settings.keys():
                    self.project_settings[tag].set("")
                    
    def generate_gcode(self):
        """Generate the GCode from the input files that have a matching configuration file
        by invoking the PCB2GCode application.
        """
        print("About to generate GCode")
        # Check inputs
        executable_path = self.plugin["ExecutablePath"]
        if not executable_path or not os.path.isfile(executable_path):
            print("Error: No PCB2GCode executable found.")
            return

        if self.project_path and os.path.isdir(self.project_path):
            output_folder = self.plugin["OutputFolder"]
            if output_folder == "":
                output_folder = "gcode"
            output_path = os.path.join(self.project_path, output_folder)
            # Create the output folder if needed
            if not os.path.isdir(output_path):
                os.mkdir(output_path)
            else: # remove any existing gcode files
                remove_files(output_path, "*.ngc")                
        else:
            print("Error: No valid output folder.")
            return
            
        joblist = []
        cmd = [executable_path, "--output-dir", output_path, "--config"]
        #Front Copper
        cfg_file = self.plugin["FrontCopperSettings"]
        in_file = os.path.join(self.project_path, self.project_settings[FRONT_COPPER].get())
        if files_exist([cfg_file, in_file]):
            joblist.append([*cmd, cfg_file, "--front", in_file])
            
        #Back Copper
        cfg_file = self.plugin["BackCopperSettings"]
        in_file = os.path.join(self.project_path, self.project_settings[BACK_COPPER].get())
        if files_exist([cfg_file, in_file]):
            joblist.append([*cmd, cfg_file, "--back", in_file])
            
        #Front Engraving
        cfg_file = self.plugin["FrontEngravingSettings"]
        in_file = os.path.join(self.project_path, self.project_settings[FRONT_ENGRAVING].get())
        if files_exist([cfg_file, in_file]):
            joblist.append([*cmd, cfg_file, "--front", in_file])
            
        #Back Engraving
        cfg_file = self.plugin["BackEngravingSettings"]
        in_file = os.path.join(self.project_path, self.project_settings[BACK_ENGRAVING].get())
        if files_exist([cfg_file, in_file]):
            joblist.append([*cmd, cfg_file, "--back", in_file])
            
        #Outline
        cfg_file = self.plugin["OutlineSettings"]
        in_file = os.path.join(self.project_path, self.project_settings[OUTLINE].get())
        if files_exist([cfg_file, in_file]):
            joblist.append([*cmd, cfg_file, "--outline", in_file])
            
        #Drilling
        cfg_file = self.plugin["DrillSettings"]
        in_file = os.path.join(self.project_path, self.project_settings[DRILLING].get())
        # Note: drill output file name is hard coded to facilitate post-processing
        if files_exist([cfg_file, in_file]):
            joblist.append([*cmd, cfg_file, "--drill", in_file, "--drill-output", DRILL_OUTPUT_FILE])
            
        # Execute all the GCode generation jobs
        print("Joblist:")
        print(joblist)
        for job in joblist:
            print("Running:", job[-1:])
            subprocess.run(job)
            
        # Perform cleanup of SVG files created as a side-effect by PCB2GCode
        remove_files(output_path, "*.svg")

        # Combine all the GCode files (may assist with visualising and verification)
        self.combine_all_gcode(output_path)
        
        # Create clones of the drill file, one for each drill tool size.
        # files are named drill_1.ngc, drill_2.ngc, etc.
        self.split_drill_sizes(output_path)
        
        # If required, do post processing on all the GCode files to remove tool changes.
        if self.project_settings[REMOVE_TOOL_CHANGES].get() == "Y":
            self.remove_tool_changes(output_path)
        
        # Then finally
        self.dismiss()
        self.app.refresh()
        self.app.setStatus(_("PCB2GCode created files in {}".format(output_path)))
              
    def split_drill_sizes(self, output_path):
        """ Splits the drill GCode file into multiple files, one per drill tool size
        """
        drill_file_path = os.path.join(output_path, DRILL_OUTPUT_FILE)
        if not os.path.isfile(drill_file_path):
            print("No drill file")
            return # nothing to do
        print("Splitting drill file by tool size...")

        # pattern matching
        tool_change_msg = re.compile(r"^\(MSG, Change tool", re.IGNORECASE)
        footer_start = re.compile(r"^G00\sZ\d.+ \( All done", re.IGNORECASE)
        header_msg1 = re.compile(r"^\( This file uses \d+ drill bit sizes", re.IGNORECASE)
        header_msg2 = re.compile(r"\( Bit sizes:", re.IGNORECASE)
        
        header_section = []
        footer_section = []
        body_sections = []
        body = []
        in_footer = False
        in_body = False
        with open(drill_file_path, "rt") as in_file:
            for gcode in in_file:
                # Check the current line for a section change
                if tool_change_msg.search(gcode):
                    if in_body: # stash the previous body section
                        body_sections.append(body)
                    in_body = True
                    body = []
                elif footer_start.search(gcode):
                    body_sections.append(body) # save the last body section
                    in_footer = True
                    in_body = False
                elif header_msg1.search(gcode):
                    gcode = "( This file uses 1 drill bit size. )\n"
                elif header_msg2.search(gcode):
                    gcode = "( Modified by post-processor )\n"
                    
                # save the current line in the relevant section
                if in_body:
                    body.append(gcode)
                elif in_footer:
                    footer_section.append(gcode)
                else:
                    header_section.append(gcode)    
        
        # Write each body section to a separate file
        tool_id = 0
        for body in body_sections:
            tool_id += 1
            file_name = DRILL_CLONE_FILE.format(tool_id)
            file_path = os.path.join(output_path, file_name)
            with open(file_path, "w") as out_file:
                print("Creating single tool drill file", file_path, "for", body[0])
                for line in header_section:
                    out_file.write(line)
                for line in body:
                    out_file.write(line)
                for line in footer_section:
                    out_file.write(line)
            
    def remove_tool_changes(self, output_path):
        """ Scans all the GCode files in the output folder and comments out
        any M6 (Tool Change) and related M0 (Pause Job) commands.
        """
        print("Removing all tool change commands...")
        gcode_files = glob.glob(os.path.join(output_path, "*.ngc"))
        if not gcode_files:
            print("No GCode files found")
            return
        # Patterns to look for
        M6_cmd = re.compile(r"^M6\s")
        M0_cmd = re.compile(r"^M0\s")
        
        # Process all the GCode files
        for ngc_file in gcode_files:
            content = []
            with open(ngc_file, "rt") as in_file:
                for gcode in in_file:
                    if M6_cmd.search(gcode) or M0_cmd.search(gcode):
                        gcode = ";" + gcode
                    content.append(gcode)
            with open(ngc_file, "wt") as out_file:
                print("Removed tool changes from", ngc_file)
                for gcode in content:
                    out_file.write(gcode)

    def combine_all_gcode(self, output_path):
        """ Scans all the GCode files in the output folder and combines them
        into a single file.
        """
        print("Combining all generated gcode...")
        gcode_files = glob.glob(os.path.join(output_path, "*.ngc"))
        if not gcode_files:
            print("No GCode files found to combine")
            return

        combined_gcode_file = os.path.join(output_path, "all.ngc")
        content = ["( Combined GCode )\n"]        
        # Process all the GCode files
        for ngc_file in gcode_files:
            with open(ngc_file, "rt") as in_file:
                for gcode in in_file:
                    content.append(gcode)
        with open(combined_gcode_file, "wt") as out_file:
            for gcode in content:
                out_file.write(gcode)
        
                    


# =============================================================================
# Create GCode from Gerber and Excellon files
# =============================================================================
class Tool(Plugin):
    __doc__ = _("Create Gcode from Gerber and Excellon files")

    def __init__(self, master):
        Plugin.__init__(self, master, "PCB2GCode")
        self.icon = "tr"
        self.group = "Development"

        self.variables = [
            ("name", "db", "", _("Name")),
            ("ExecutablePath", "file", "", _("Path to PCB2GCode executable")),
            ("FrontCopperSettings", "file", "", _("Front copper settings file")),
            ("BackCopperSettings", "file", "", _("Back copper settings file")),
            ("FrontEngravingSettings", "file", "", _("Front engraving settings file")),
            ("BackEngravingSettings", "file", "", _("Back engraving settings file")),
            ("OutlineSettings", "file", "", _("Board outline settings file")),
            ("DrillSettings", "file", "", _("Drill settings file")),
            ("ProjectPath", "text", "", _("Project path")),
            ("OutputFolder", "text", "gcode", _("Output folder")),
        ]
        self.help = "\n".join([
            "This plugin generates GCode from Gerber and Excellon files using the PCB2Code utility.",
            "",
        ])

        self.buttons.append("exe")


    # ----------------------------------------------------------------------
    def execute(self, app):
        """Invoke the PCB2GCode form to handle the user interaction and conversion"""
        dlg = PCB2GCode(self, app)
        dlg.show()       

      
    


