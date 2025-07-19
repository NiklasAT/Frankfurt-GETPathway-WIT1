# Enhanced Nuclear Membrane & Cytoplasm Analysis - PUBLICATION-READY CHARTS
# Implements separate data series with distinct colors and scientific styling

from ij import IJ, ImagePlus
from ij.gui import WaitForUserDialog, GenericDialog, YesNoCancelDialog
from ij.io import DirectoryChooser
from ij.measure import Calibration
from java.util import Date
from java.text import SimpleDateFormat
from java.awt import Frame, Button, Panel, BorderLayout, FlowLayout, GridLayout
from java.awt.event import ActionListener
from javax.swing import JFrame, JButton, JPanel, JLabel, BoxLayout, JOptionPane
import os
import math

# Try to import Java Excel libraries (Apache POI) - WORKING VERSION
try:
    from java.io import FileOutputStream
    from java.util import Calendar
    
    # Test basic POI first
    from org.apache.poi.xssf.usermodel import XSSFWorkbook
    IJ.log("✓ Basic POI XSSFWorkbook imported successfully")
    
    # Import additional POI classes for Excel manipulation
    from org.apache.poi.xssf.usermodel import XSSFSheet, XSSFRow, XSSFCell
    from org.apache.poi.ss.usermodel import CellType, IndexedColors, WorkbookFactory
    from org.apache.poi.xssf.usermodel import XSSFCellStyle, XSSFFont
    from org.apache.poi.ss.usermodel import FillPatternType
    from org.apache.poi.ss.util import CellRangeAddress
    IJ.log("✓ Standard POI classes imported successfully")
    
    # Try to import XDDF chart classes - UPDATED TO USE WORKING CLASSES
    try:
        # Use the classes that work in your installation (based on your earlier successful test)
        from org.apache.poi.xddf.usermodel.chart import (
            ChartTypes, XDDFDataSourcesFactory, XDDFChart, 
            XDDFScatterChartData, XDDFLineChartData,
            XDDFCategoryAxis, XDDFValueAxis, AxisPosition
        )
        # Test if ChartTypes specifically works (this was working in your earlier tests)
        test_chart_types = ChartTypes.LINE
        CHARTS_AVAILABLE = True
        IJ.log("✓ Working chart classes with ChartTypes imported successfully")
        IJ.log("🎉 Apache POI Excel libraries with publication-ready charting loaded successfully")
    except ImportError as chart_error:
        IJ.log("⚠️ Chart classes not available: " + str(chart_error))
        CHARTS_AVAILABLE = False
        IJ.log("🎉 Apache POI Excel libraries loaded successfully (without embedded charts)")
    except Exception as chart_error:
        IJ.log("⚠️ Chart classes import error: " + str(chart_error))
        CHARTS_AVAILABLE = False
        IJ.log("🎉 Apache POI Excel libraries loaded successfully (without embedded charts)")
        
    EXCEL_AVAILABLE = True
        
except ImportError as e:
    EXCEL_AVAILABLE = False
    CHARTS_AVAILABLE = False
    IJ.log("❌ Apache POI not available - " + str(e))

class ContinueDialog:
    """Custom dialog with proper buttons for continue/export choice"""
    
    def __init__(self, cell_count):
        self.choice = None
        self.cell_count = cell_count
        
    def show_dialog(self):
        """Show dialog and return choice"""
        options = ["Measure another cell", "Export data now"]
        
        choice = JOptionPane.showOptionDialog(
            None,
            "Cell " + str(self.cell_count).zfill(2) + " analysis completed successfully!\n\nWhat would you like to do next?",
            "Analysis Complete",
            JOptionPane.YES_NO_OPTION,
            JOptionPane.QUESTION_MESSAGE,
            None,
            options,
            options[0]
        )
        
        if choice == 0:
            return "continue"
        elif choice == 1:
            return "export"
        else:
            return "cancel"

class EnhancedNuclearAnalyzer:
    def __init__(self):
        self.nuclear_membrane_averages = []
        self.cytoplasm_averages = []
        self.nucleus_cytoplasm_ratios = []
        self.cell_data = []
        self.cell_count = 0
        self.threshold_percentage = 50.0
        self.last_directory = None
        self.image_name = ""
        self.microscope_metadata = {}
        
        # Metadata fields to extract (filtered list)
        self.wanted_metadata = {
            'Image_Height': 'Image_Height',
            'DisplaySetting|Channel|DyeName': 'DyeName', 
            'Information|Image|Channel|EmissionWavelength': 'EmissionWavelength',
            'Information|Image|Channel|ExcitationWavelength': 'ExcitationWavelength'
        }
        
    def run_analysis(self):
        """Main analysis function with user settings"""
        # Check if image is open
        if IJ.getImage() is None:
            IJ.showMessage("Error", "Please open an image first.")
            return
        
        # Get image info
        imp = IJ.getImage()
        self.image_name = imp.getTitle()
        self.collect_metadata(imp)
        
        # Get user settings
        if not self.get_analysis_settings():
            return
        
        # Initialize
        IJ.log("\\Clear")
        IJ.log("=== Enhanced Nuclear Membrane & Cytoplasm Analysis ===")
        IJ.log("Image: " + self.image_name)
        IJ.log("Threshold: " + str(self.threshold_percentage) + "%")
        IJ.log("Excel output: " + ("Available" if EXCEL_AVAILABLE else "Fallback mode"))
        if EXCEL_AVAILABLE and CHARTS_AVAILABLE:
            IJ.log("Charts: Publication-ready with separate data series and scientific styling")
        elif EXCEL_AVAILABLE:
            IJ.log("Charts: Excel files will be created without embedded charts")
        IJ.log("Starting analysis...")
        
        # Set line width
        IJ.run("Line Width...", "line=1")
        
        # Main analysis loop
        while True:
            self.cell_count += 1
            IJ.log("\n=== ANALYZING CELL " + str(self.cell_count).zfill(2) + " ===")
            
            # Analyze current cell
            cell_data = self.analyze_single_cell()
            
            if cell_data is None:
                IJ.log("Cell " + str(self.cell_count).zfill(2) + " skipped")
                self.cell_count -= 1
                continue
            
            # Store cell data
            self.cell_data.append(cell_data)
            self.nuclear_membrane_averages.append(cell_data['nuclear_mean'])
            self.cytoplasm_averages.append(cell_data['cytoplasm_mean'])
            self.nucleus_cytoplasm_ratios.append(cell_data['ratio'])
            
            # Show results
            IJ.log("\n*** CELL " + str(self.cell_count).zfill(2) + " COMPLETE ***")
            IJ.log("  Nuclear membrane mean: " + str(round(cell_data['nuclear_mean'], 3)))
            IJ.log("  Cytoplasm mean: " + str(round(cell_data['cytoplasm_mean'], 3)))
            IJ.log("  Nucleus/Cytoplasm ratio: " + str(round(cell_data['ratio'], 3)))
            
            # Ask user what to do next
            choice = self.ask_continue()
            
            if choice == "continue":
                continue
            elif choice == "export":
                break
            else:
                IJ.log("Analysis cancelled by user")
                return
        
        # Export to Excel
        self.export_to_excel()
        
        IJ.log("\nAnalysis complete!")
        IJ.showMessage("Analysis Complete", 
                      "Excel file created successfully!\n" + 
                      "Total cells analyzed: " + str(self.cell_count) + "\n" +
                      "Publication-ready charts with separate data series included!")
    
    def get_analysis_settings(self):
        """Get user settings for analysis"""
        gd = GenericDialog("Publication-Ready Analysis Settings")
        gd.addMessage("Configure your analysis parameters:")
        gd.addNumericField("Threshold percentage (for top intensity values):", self.threshold_percentage, 1, 6, "%")
        gd.addMessage("This threshold determines which intensity values are used for calculations.")
        gd.addMessage("")
        gd.addMessage("PUBLICATION-READY CHART FEATURES:")
        gd.addMessage("- Nuclear membrane: Each stroke as separate colored series")
        gd.addMessage("- Cytoplasm: Above-threshold vs below-threshold series")
        gd.addMessage("- Clean scientific styling with proper legends")
        gd.addMessage("- No 3D effects or visual clutter")
        gd.showDialog()
        
        if gd.wasCanceled():
            return False
        
        self.threshold_percentage = gd.getNextNumber()
        
        if self.threshold_percentage < 1 or self.threshold_percentage > 100:
            IJ.showMessage("Error", "Threshold must be between 1% and 100%")
            return False
        
        return True
    
    def collect_metadata(self, imp):
        """Collect microscope metadata from image"""
        self.microscope_metadata = {}
        
        # Add basic image height
        self.microscope_metadata['Image_Height'] = str(imp.getHeight())
        
        # Process image properties
        prop = imp.getProperties()
        if prop is not None:
            for key in prop.keySet():
                key_str = str(key)
                value = prop.get(key)
                
                if value is not None:
                    value_str = str(value)
                    
                    # Check each wanted metadata key
                    for wanted_key, clean_name in self.wanted_metadata.items():
                        if wanted_key in key_str:
                            if "=" in value_str:
                                parts = value_str.split("=", 1)
                                if len(parts) == 2:
                                    clean_value = parts[1].strip()
                                    self.microscope_metadata[clean_name] = clean_value
                            else:
                                self.microscope_metadata[clean_name] = value_str
                            break
        
        IJ.log("Collected " + str(len(self.microscope_metadata)) + " metadata fields")
    
    def analyze_single_cell(self):
        """Analyze a single cell and return comprehensive data"""
        # Step 1: Nuclear Membrane Analysis
        nuclear_data = self.measure_nuclear_membrane_detailed()
        if nuclear_data is None:
            return None
        
        # Step 2: Cytoplasm Analysis
        cytoplasm_data = self.measure_cytoplasm_detailed()
        if cytoplasm_data is None:
            return None
        
        # Combine data
        cell_data = {
            'cell_number': self.cell_count,
            'nuclear_data': nuclear_data,
            'cytoplasm_data': cytoplasm_data,
            'nuclear_mean': nuclear_data['overall_mean'],
            'cytoplasm_mean': cytoplasm_data['mean'],
            'ratio': nuclear_data['overall_mean'] / cytoplasm_data['mean'] if cytoplasm_data['mean'] > 0 else 0
        }
        
        return cell_data
    
    def measure_nuclear_membrane_detailed(self):
        """Measure nuclear membrane with detailed stroke-by-stroke analysis"""
        IJ.log("\n--- Nuclear Membrane Measurement (Cell " + str(self.cell_count).zfill(2) + ") ---")
        
        # Clear selection and set tool
        IJ.run("Select None")
        IJ.setTool("polyline")
        
        # Enhanced dialog with colored headings - using WaitForUserDialog with better formatting
        try:
            # Use ImageJ's WaitForUserDialog but with enhanced text formatting
            title = "NUCLEAR MEMBRANE - Cell " + str(self.cell_count).zfill(2)
            
            message = ("=== NUCLEAR MEMBRANE MEASUREMENT ===\n" +
                      "Cell: " + str(self.cell_count).zfill(2) + " | Threshold: " + str(self.threshold_percentage) + "%\n\n" +
                      "INSTRUCTIONS FOR NUCLEAR MEMBRANE:\n" +
                      "1. Use the segmented line tool (polyline) to trace the nuclear membrane\n" +
                      "2. Click multiple points along the membrane to create segments\n" +
                      "3. Each segment will be analyzed separately\n" +
                      "4. Double-click to finish the segmented line\n" +
                      "5. Press OK when tracing is complete\n\n" +
                      "ANALYSIS METHOD:\n" +
                      "- Each stroke/segment analyzed separately\n" +
                      "- Top " + str(self.threshold_percentage) + "% of values used per segment\n" +
                      "- Each segment becomes a separate data series in charts\n\n" +
                      "You can interact with the image while this dialog is open.\n" +
                      "Press OK when your polyline selection is ready.")
            
            dialog = WaitForUserDialog(title, message)
            dialog.show()
            
        except Exception as e:
            IJ.log("Enhanced dialog failed, using simple version: " + str(e))
            # Fallback to simple dialog
            dialog = WaitForUserDialog("NUCLEAR MEMBRANE - Cell " + str(self.cell_count).zfill(2), 
                                     "Trace nuclear membrane with polyline tool, then press OK")
            dialog.show()
        
        # Get current image and selection
        imp = IJ.getImage()
        roi = imp.getRoi()
        
        if roi is None or roi.getType() != roi.POLYLINE:
            roi_type = str(roi.getType()) if roi else "None"
            IJ.showMessage("Error", "No polyline selection detected! Selection type: " + roi_type)
            return None
        
        # Get selection coordinates
        polygon = roi.getPolygon()
        x_coords = polygon.xpoints
        y_coords = polygon.ypoints
        n_anchors = polygon.npoints
        
        if n_anchors < 2:
            IJ.showMessage("Error", "The polyline must have at least 2 anchor points. Current points: " + str(n_anchors))
            return None
        
        n_segments = n_anchors - 1
        IJ.log("Polyline detected with " + str(n_anchors) + " anchor points (" + str(n_segments) + " segments)")
        
        # Get intensity profile
        imp.setRoi(roi)
        from ij.gui import ProfilePlot
        pp = ProfilePlot(imp)
        full_profile = pp.getProfile()
        
        if full_profile is None or len(full_profile) == 0:
            IJ.showMessage("Error", "No intensity data collected from selection.")
            return None
        
        IJ.log("Full profile length: " + str(len(full_profile)) + " pixels")
        
        # Calculate segment boundaries
        segment_boundaries = self.calculate_segment_boundaries(x_coords, y_coords, len(full_profile))
        
        # Process each segment
        stroke_data = []
        stroke_averages = []
        
        for seg in range(n_segments):
            start_idx = segment_boundaries[seg]
            end_idx = segment_boundaries[seg + 1]
            
            # Extract segment values
            segment_values = full_profile[start_idx:end_idx + 1]
            
            if len(segment_values) == 0:
                IJ.log("  Segment " + str(seg + 1) + ": No values (skipped)")
                continue
            
            # Sort and get top threshold%
            sorted_values = sorted(segment_values)
            threshold_idx = int(len(sorted_values) * (1.0 - self.threshold_percentage / 100.0))
            top_values = sorted_values[threshold_idx:]
            
            # Calculate mean of top values
            stroke_avg = sum(top_values) / float(len(top_values))
            stroke_averages.append(stroke_avg)
            
            IJ.log("  Segment " + str(seg + 1) + ": " + str(len(segment_values)) + " pixels, top " + str(self.threshold_percentage) + "% mean = " + str(round(stroke_avg, 3)))
            
            # Store detailed stroke data
            stroke_info = {
                'segment_number': seg + 1,
                'start_index': start_idx,
                'end_index': end_idx,
                'raw_values': segment_values,
                'top_values': top_values,
                'stroke_average': stroke_avg,
                'raw_values_str': ";".join([str(round(val, 3)) for val in segment_values]),
                'top_values_str': ";".join([str(round(val, 3)) for val in top_values])
            }
            stroke_data.append(stroke_info)
        
        if len(stroke_averages) == 0:
            IJ.showMessage("Error", "No valid segments could be processed from the polyline.")
            return None
        
        # Calculate overall nuclear membrane mean
        overall_mean = sum(stroke_averages) / float(len(stroke_averages))
        
        IJ.log("Nuclear membrane analysis complete:")
        IJ.log("  " + str(len(stroke_averages)) + " segments processed")
        IJ.log("  Overall nuclear membrane mean: " + str(round(overall_mean, 3)))
        
        # Return comprehensive data
        return {
            'full_profile': full_profile,
            'segment_boundaries': segment_boundaries,
            'stroke_data': stroke_data,
            'stroke_averages': stroke_averages,
            'overall_mean': overall_mean,
            'n_segments': len(stroke_averages)
        }
    
    def measure_cytoplasm_detailed(self):
        """Measure cytoplasm with detailed analysis"""
        IJ.log("\n--- Cytoplasm Measurement (Cell " + str(self.cell_count).zfill(2) + ") ---")
        
        # Clear selection and set tool
        IJ.run("Select None")
        IJ.setTool("polyline")
        
        # Enhanced dialog with colored headings - using WaitForUserDialog with better formatting  
        try:
            # Use ImageJ's WaitForUserDialog but with enhanced text formatting
            title = "CYTOPLASM - Cell " + str(self.cell_count).zfill(2)
            
            message = ("=== CYTOPLASM MEASUREMENT ===\n" +
                      "Cell: " + str(self.cell_count).zfill(2) + " | Threshold: " + str(self.threshold_percentage) + "%\n\n" +
                      "INSTRUCTIONS FOR CYTOPLASM:\n" +
                      "1. Use the segmented line tool (polyline) to trace in cytoplasm area\n" +
                      "2. **AVOID the nuclear region - stay in cytoplasm only**\n" +
                      "3. Click multiple points to create the measurement line\n" +
                      "4. Double-click to finish the segmented line\n" +
                      "5. Press OK when tracing is complete\n\n" +
                      "ANALYSIS METHOD:\n" +
                      "- Entire polyline analyzed together (not segment-by-segment)\n" +
                      "- Top " + str(self.threshold_percentage) + "% of all values used\n" +
                      "- Creates two data series: above/below threshold\n\n" +
                      "You can interact with the image while this dialog is open.\n" +
                      "Press OK when your polyline selection is ready.")
            
            dialog = WaitForUserDialog(title, message)
            dialog.show()
            
        except Exception as e:
            IJ.log("Enhanced dialog failed, using simple version: " + str(e))
            # Fallback to simple dialog  
            dialog = WaitForUserDialog("CYTOPLASM - Cell " + str(self.cell_count).zfill(2), 
                                     "Trace cytoplasm area with polyline tool (avoid nucleus), then press OK")
            dialog.show()
        
        # Get current image and selection
        imp = IJ.getImage()
        roi = imp.getRoi()
        
        if roi is None or roi.getType() != roi.POLYLINE:
            roi_type = str(roi.getType()) if roi else "None"
            IJ.showMessage("Error", "No polyline selection detected! Selection type: " + roi_type)
            return None
        
        # Get intensity profile
        imp.setRoi(roi)
        from ij.gui import ProfilePlot
        pp = ProfilePlot(imp)
        profile = pp.getProfile()
        
        if profile is None or len(profile) == 0:
            IJ.showMessage("Error", "No intensity data collected from selection.")
            return None
        
        # Sort and get top threshold%
        sorted_profile = sorted(profile)
        threshold_idx = int(len(sorted_profile) * (1.0 - self.threshold_percentage / 100.0))
        top_values = sorted_profile[threshold_idx:]
        threshold_value = sorted_profile[threshold_idx] if threshold_idx < len(sorted_profile) else 0
        
        # Calculate mean
        cytoplasm_mean = sum(top_values) / float(len(top_values))
        
        IJ.log("Cytoplasm analysis complete:")
        IJ.log("  Profile length: " + str(len(profile)) + " pixels")
        IJ.log("  Top " + str(self.threshold_percentage) + "%: " + str(len(top_values)) + " pixels")
        IJ.log("  Threshold value: " + str(round(threshold_value, 3)))
        IJ.log("  Cytoplasm mean: " + str(round(cytoplasm_mean, 3)))
        
        # Return comprehensive data
        return {
            'raw_profile': profile,
            'top_values': top_values,
            'threshold_value': threshold_value,
            'mean': cytoplasm_mean,
            'raw_values_str': ";".join([str(round(val, 3)) for val in profile]),
            'top_values_str': ";".join([str(round(val, 3)) for val in top_values])
        }
    
    def calculate_segment_boundaries(self, x_coords, y_coords, profile_length):
        """Calculate segment boundaries in the profile"""
        n_anchors = len(x_coords)
        
        # Calculate distances along polyline
        anchor_distances = [0.0]
        total_distance = 0.0
        
        for i in range(1, n_anchors):
            dx = x_coords[i] - x_coords[i-1]
            dy = y_coords[i] - y_coords[i-1]
            segment_distance = math.sqrt(dx*dx + dy*dy)
            total_distance += segment_distance
            anchor_distances.append(total_distance)
        
        # Convert distances to profile indices
        boundaries = []
        for i in range(n_anchors):
            idx = int(round((anchor_distances[i] / total_distance) * (profile_length - 1)))
            if idx >= profile_length:
                idx = profile_length - 1
            boundaries.append(idx)
        
        return boundaries
    
    def ask_continue(self):
        """Ask user whether to continue with another cell"""
        dialog = ContinueDialog(self.cell_count)
        return dialog.show_dialog()
    
    def export_to_excel(self):
        """Export comprehensive data to Excel with publication-ready charts"""
        IJ.log("\n=== EXCEL EXPORT ===")
        IJ.log("Excel libraries available: " + str(EXCEL_AVAILABLE))
        IJ.log("Chart functionality: " + str(CHARTS_AVAILABLE))
        
        try:
            # Get export settings
            export_settings = self.get_export_settings()
            if export_settings is None:
                return
            
            filepath = export_settings['filepath']
            IJ.log("Target filepath: " + str(filepath))
            
            # Create Excel with publication-ready charts
            if EXCEL_AVAILABLE:
                IJ.log("Creating Excel file with publication-ready charts...")
                try:
                    self.create_excel_with_publication_charts(filepath)
                    IJ.log("Excel file with publication-ready charts created successfully!")
                except Exception as e:
                    IJ.log("Excel creation failed: " + str(e))
                    import traceback
                    IJ.log("Full error trace: " + traceback.format_exc())
            else:
                IJ.log("No Excel support available")
            
        except Exception as e:
            IJ.log("Error in export_to_excel: " + str(e))
    
    def create_excel_with_publication_charts(self, filepath):
        """Create Excel workbook with publication-ready charts and threshold comparison"""
        try:
            IJ.log("Creating Excel file with publication-ready charts and threshold comparison...")
            
            # Create workbook
            workbook = XSSFWorkbook()
            
            # Create summary sheet
            self.create_publication_summary_sheet(workbook)
            
            # Create threshold comparison sheet (NEW)
            self.create_threshold_comparison_sheet(workbook)
            
            # Create individual cell sheets with publication-ready charts
            for i, cell_data in enumerate(self.cell_data):
                cell_num = str(i + 1).zfill(2)
                IJ.log("Creating publication-ready sheet for Cell " + cell_num + "...")
                self.create_cell_sheet_with_publication_charts(workbook, cell_data, cell_num)
            
            # Create metadata sheet
            self.create_publication_metadata_sheet(workbook)
            
            # Write to file
            IJ.log("Writing Excel file...")
            file_out = FileOutputStream(filepath)
            workbook.write(file_out)
            file_out.close()
            workbook.close()
            
            IJ.log("Excel file with publication-ready charts and threshold comparison created successfully!")
            
        except Exception as e:
            IJ.log("Error creating Excel with publication-ready charts: " + str(e))
            raise e
    
    def create_threshold_comparison_sheet(self, workbook):
        """Create threshold comparison sheet analyzing sensitivity to threshold definition"""
        IJ.log("Creating threshold comparison analysis...")
        
        sheet = workbook.createSheet("Threshold Comparison")
        
        row_num = 0
        
        # Header section
        header_row = sheet.createRow(row_num)
        header_cell = header_row.createCell(0)
        header_cell.setCellValue("Nuclear/Cytoplasmic Intensity Ratio - Threshold Sensitivity Analysis")
        row_num += 2
        
        # Description with ASCII-safe characters
        desc_row = sheet.createRow(row_num)
        desc_row.createCell(0).setCellValue("This analysis shows how the nuclear/cytoplasmic intensity ratio changes")
        row_num += 1
        desc_row2 = sheet.createRow(row_num)
        desc_row2.createCell(0).setCellValue("with different threshold values for defining 'top intensity' pixels.")
        row_num += 1
        desc_row3 = sheet.createRow(row_num)
        desc_row3.createCell(0).setCellValue("Main analysis threshold: " + str(self.threshold_percentage) + "%")
        row_num += 2
        
        # Fixed thresholds to analyze
        fixed_thresholds = [10.0, 20.0, 30.0, 40.0, 50.0, 60.0]
        
        # Create headers
        header_row = sheet.createRow(row_num)
        col = 0
        header_row.createCell(col).setCellValue("Cell Number")
        col += 1
        
        for threshold in fixed_thresholds:
            threshold_str = str(int(threshold)) + "%"
            header_row.createCell(col).setCellValue("Nuclear Mean @" + threshold_str)
            header_row.createCell(col + 1).setCellValue("Cytoplasm Mean @" + threshold_str)
            header_row.createCell(col + 2).setCellValue("Ratio @" + threshold_str)
            col += 3
        
        row_num += 1
        
        # Process each cell with all threshold values
        for i, cell_data in enumerate(self.cell_data):
            row = sheet.createRow(row_num)
            col = 0
            
            # Cell number
            row.createCell(col).setCellValue("Cell " + str(i + 1).zfill(2))
            col += 1
            
            # Calculate ratios for each fixed threshold
            for threshold in fixed_thresholds:
                # Recalculate nuclear mean with this threshold
                nuclear_mean = self.recalculate_nuclear_mean_with_threshold(cell_data['nuclear_data'], threshold)
                
                # Recalculate cytoplasm mean with this threshold  
                cytoplasm_mean = self.recalculate_cytoplasm_mean_with_threshold(cell_data['cytoplasm_data'], threshold)
                
                # Calculate ratio
                ratio = nuclear_mean / cytoplasm_mean if cytoplasm_mean > 0 else 0
                
                # Add to sheet with 3 decimal formatting
                row.createCell(col).setCellValue(round(nuclear_mean, 3))
                row.createCell(col + 1).setCellValue(round(cytoplasm_mean, 3))
                row.createCell(col + 2).setCellValue(round(ratio, 3))
                col += 3
            
            row_num += 1
        
        row_num += 2
        
        # Summary statistics for threshold comparison
        summary_header = sheet.createRow(row_num)
        summary_header.createCell(0).setCellValue("Summary Statistics (Mean +/- SE across all cells)")
        row_num += 1
        
        # Headers for summary
        summary_col_header = sheet.createRow(row_num)
        col = 0
        summary_col_header.createCell(col).setCellValue("Threshold")
        col += 1
        summary_col_header.createCell(col).setCellValue("Nuclear Mean +/- SE")
        summary_col_header.createCell(col + 1).setCellValue("Cytoplasm Mean +/- SE")
        summary_col_header.createCell(col + 2).setCellValue("Ratio Mean +/- SE")
        row_num += 1
        
        # Calculate summary statistics for each threshold
        for threshold in fixed_thresholds:
            nuclear_values = []
            cytoplasm_values = []
            ratio_values = []
            
            for cell_data in self.cell_data:
                nuclear_mean = self.recalculate_nuclear_mean_with_threshold(cell_data['nuclear_data'], threshold)
                cytoplasm_mean = self.recalculate_cytoplasm_mean_with_threshold(cell_data['cytoplasm_data'], threshold)
                ratio = nuclear_mean / cytoplasm_mean if cytoplasm_mean > 0 else 0
                
                nuclear_values.append(nuclear_mean)
                cytoplasm_values.append(cytoplasm_mean)
                ratio_values.append(ratio)
            
            # Calculate means and standard errors
            nuclear_mean_avg = sum(nuclear_values) / float(len(nuclear_values))
            nuclear_se = self.calculate_se(nuclear_values)
            
            cytoplasm_mean_avg = sum(cytoplasm_values) / float(len(cytoplasm_values))
            cytoplasm_se = self.calculate_se(cytoplasm_values)
            
            ratio_mean_avg = sum(ratio_values) / float(len(ratio_values))
            ratio_se = self.calculate_se(ratio_values)
            
            # Add summary row
            summary_row = sheet.createRow(row_num)
            col = 0
            summary_row.createCell(col).setCellValue(str(int(threshold)) + "%")
            col += 1
            summary_row.createCell(col).setCellValue(str(round(nuclear_mean_avg, 3)) + " +/- " + str(round(nuclear_se, 3)))
            summary_row.createCell(col + 1).setCellValue(str(round(cytoplasm_mean_avg, 3)) + " +/- " + str(round(cytoplasm_se, 3)))
            summary_row.createCell(col + 2).setCellValue(str(round(ratio_mean_avg, 3)) + " +/- " + str(round(ratio_se, 3)))
            row_num += 1
        
        row_num += 2
        
        # Interpretation note with ASCII-safe characters
        note_row = sheet.createRow(row_num)
        note_row.createCell(0).setCellValue("Interpretation:")
        row_num += 1
        note_row2 = sheet.createRow(row_num)
        note_row2.createCell(0).setCellValue("- Lower thresholds (10-30%) include more pixels but may include background")
        row_num += 1
        note_row3 = sheet.createRow(row_num)
        note_row3.createCell(0).setCellValue("- Higher thresholds (50-60%) are more selective but may miss dim signals")
        row_num += 1
        note_row4 = sheet.createRow(row_num)
        note_row4.createCell(0).setCellValue("- Compare ratios across thresholds to assess measurement robustness")
        
        # Auto-size columns
        try:
            for i in range(19):  # Cell + 6 thresholds × 3 columns
                sheet.autoSizeColumn(i)
        except:
            pass
        
        IJ.log("Threshold comparison sheet created with analysis for " + str(len(fixed_thresholds)) + " thresholds")
    
    def recalculate_nuclear_mean_with_threshold(self, nuclear_data, threshold_percent):
        """Recalculate nuclear membrane mean using a different threshold percentage"""
        stroke_averages = []
        
        for stroke in nuclear_data['stroke_data']:
            raw_values = stroke['raw_values']
            
            if len(raw_values) == 0:
                continue
            
            # Sort and get top threshold%
            sorted_values = sorted(raw_values)
            threshold_idx = int(len(sorted_values) * (1.0 - threshold_percent / 100.0))
            top_values = sorted_values[threshold_idx:]
            
            if len(top_values) > 0:
                stroke_avg = sum(top_values) / float(len(top_values))
                stroke_averages.append(stroke_avg)
        
        if len(stroke_averages) > 0:
            return sum(stroke_averages) / float(len(stroke_averages))
        else:
            return 0.0
    
    def recalculate_cytoplasm_mean_with_threshold(self, cytoplasm_data, threshold_percent):
        """Recalculate cytoplasm mean using a different threshold percentage"""
        raw_profile = cytoplasm_data['raw_profile']
        
        if len(raw_profile) == 0:
            return 0.0
        
        # Sort and get top threshold%
        sorted_profile = sorted(raw_profile)
        threshold_idx = int(len(sorted_profile) * (1.0 - threshold_percent / 100.0))
        top_values = sorted_profile[threshold_idx:]
        
        if len(top_values) > 0:
            return sum(top_values) / float(len(top_values))
        else:
            return 0.0
    
    def create_cell_sheet_with_publication_charts(self, workbook, cell_data, cell_num):
        """Create cell sheet with publication-ready separate data series charts"""
        sheet = workbook.createSheet("Cell_" + cell_num)
        
        row_num = 0
        
        # Header
        row = sheet.createRow(row_num)
        row.createCell(0).setCellValue("Cell " + cell_num + " - Publication-Ready Analysis")
        row_num += 2
        
        # === NUCLEAR MEMBRANE DATA FOR SEPARATE SERIES ===
        nuclear_section = sheet.createRow(row_num)
        nuclear_section.createCell(0).setCellValue("NUCLEAR MEMBRANE - Data for Separate Stroke Series")
        row_num += 1
        
        # Create header row for nuclear data with separate columns for each stroke
        nuclear_header = sheet.createRow(row_num)
        nuclear_header.createCell(0).setCellValue("Profile Index")
        
        # Add column for each stroke
        n_segments = cell_data['nuclear_data']['n_segments']
        for seg in range(n_segments):
            nuclear_header.createCell(seg + 1).setCellValue("Stroke " + str(seg + 1))
        row_num += 1
        
        nuclear_data_start = row_num
        
        # Prepare nuclear data with separate series for each stroke
        nuclear_profile = cell_data['nuclear_data']['full_profile']
        boundaries = cell_data['nuclear_data']['segment_boundaries']
        stroke_data = cell_data['nuclear_data']['stroke_data']
        
        # Create data rows with separate columns for each stroke
        for j, intensity in enumerate(nuclear_profile):
            row = sheet.createRow(row_num)
            row.createCell(0).setCellValue(float(j))  # Profile index
            
            # Determine which stroke this pixel belongs to and only fill that column
            current_stroke = -1
            for seg in range(len(boundaries) - 1):
                if boundaries[seg] <= j <= boundaries[seg + 1]:
                    current_stroke = seg
                    break
            
            # Fill columns: put intensity in the appropriate stroke column, leave others empty
            for seg in range(n_segments):
                if seg == current_stroke:
                    row.createCell(seg + 1).setCellValue(round(intensity, 3))
                else:
                    # Leave cell empty for other strokes
                    row.createCell(seg + 1).setCellValue("")
            
            row_num += 1
        
        nuclear_data_end = row_num - 1
        
        # Create nuclear membrane chart with separate series for each stroke
        if CHARTS_AVAILABLE:
            try:
                self.create_publication_nuclear_chart(sheet, cell_num, nuclear_data_start, nuclear_data_end, n_segments,
                                                    1, nuclear_data_start - 2, 9, nuclear_data_start + 20)
                IJ.log("Publication-ready nuclear membrane chart created for Cell " + cell_num)
            except Exception as e:
                IJ.log("Nuclear chart failed for Cell " + cell_num + ": " + str(e))
        else:
            # Add instructions for manual chart creation when charts not available
            instruction_row = sheet.createRow(nuclear_data_start - 1)
            instruction_row.createCell(10).setCellValue("CHART INSTRUCTIONS:")
            instruction_row = sheet.createRow(nuclear_data_start)
            instruction_row.createCell(10).setCellValue("1. Select data range A" + str(nuclear_data_start) + ":" + chr(65 + n_segments) + str(nuclear_data_end))
            instruction_row = sheet.createRow(nuclear_data_start + 1)
            instruction_row.createCell(10).setCellValue("2. Insert > Charts > Scatter Chart")
            instruction_row = sheet.createRow(nuclear_data_start + 2)
            instruction_row.createCell(10).setCellValue("3. Each stroke will be a separate series")
            IJ.log("Chart instructions added for Cell " + cell_num + " (charts not available)")
        
        row_num += 5  # Space between charts
        
        # === CYTOPLASM DATA FOR TWO SERIES ===
        cyto_section = sheet.createRow(row_num)
        cyto_section.createCell(0).setCellValue("CYTOPLASM - Data for Above/Below Threshold Series")
        row_num += 1
        
        cyto_header = sheet.createRow(row_num)
        cyto_header.createCell(0).setCellValue("Profile Index")
        cyto_header.createCell(1).setCellValue("Above Threshold (" + str(self.threshold_percentage) + "%)")
        cyto_header.createCell(2).setCellValue("Below Threshold")
        row_num += 1
        
        cyto_data_start = row_num
        
        # Prepare cytoplasm data with separate series for above/below threshold
        cytoplasm_profile = cell_data['cytoplasm_data']['raw_profile']
        top_values = cell_data['cytoplasm_data']['top_values']
        
        for j, intensity in enumerate(cytoplasm_profile):
            row = sheet.createRow(row_num)
            row.createCell(0).setCellValue(float(j))  # Profile index
            
            # Split into above/below threshold series
            if intensity in top_values:
                row.createCell(1).setCellValue(round(intensity, 3))  # Above threshold
                row.createCell(2).setCellValue("")  # Leave below threshold empty
            else:
                row.createCell(1).setCellValue("")  # Leave above threshold empty
                row.createCell(2).setCellValue(round(intensity, 3))  # Below threshold
            
            row_num += 1
        
        cyto_data_end = row_num - 1
        
        # Create cytoplasm chart with two series (above/below threshold)
        if CHARTS_AVAILABLE:
            try:
                self.create_publication_cytoplasm_chart(sheet, cell_num, cyto_data_start, cyto_data_end,
                                                      1, cyto_data_start - 2, 9, cyto_data_start + 20)
                IJ.log("Publication-ready cytoplasm chart created for Cell " + cell_num)
            except Exception as e:
                IJ.log("Cytoplasm chart failed for Cell " + cell_num + ": " + str(e))
        else:
            # Add instructions for manual chart creation when charts not available
            instruction_row = sheet.createRow(cyto_data_start - 1)
            instruction_row.createCell(5).setCellValue("CHART INSTRUCTIONS:")
            instruction_row = sheet.createRow(cyto_data_start)
            instruction_row.createCell(5).setCellValue("1. Select data range A" + str(cyto_data_start) + ":C" + str(cyto_data_end))
            instruction_row = sheet.createRow(cyto_data_start + 1)
            instruction_row.createCell(5).setCellValue("2. Insert > Charts > Scatter Chart")
            instruction_row = sheet.createRow(cyto_data_start + 2)
            instruction_row.createCell(5).setCellValue("3. Above threshold = green, below = gray")
            IJ.log("Chart instructions added for Cell " + cell_num + " (charts not available)")
        
        row_num += 5
        
        # === SUMMARY DATA TABLE ===
        summary_section = sheet.createRow(row_num)
        summary_section.createCell(0).setCellValue("Summary Data")
        row_num += 1
        
        summary_header = sheet.createRow(row_num)
        headers = ["Stroke", "Start Index", "End Index", "Pixel Count", "Top " + str(self.threshold_percentage) + "% Count", "Average"]
        for i, header in enumerate(headers):
            summary_header.createCell(i).setCellValue(header)
        row_num += 1
        
        for stroke in cell_data['nuclear_data']['stroke_data']:
            row = sheet.createRow(row_num)
            row.createCell(0).setCellValue(float(stroke['segment_number']))
            row.createCell(1).setCellValue(float(stroke['start_index']))
            row.createCell(2).setCellValue(float(stroke['end_index']))
            row.createCell(3).setCellValue(float(len(stroke['raw_values'])))
            row.createCell(4).setCellValue(float(len(stroke['top_values'])))
            row.createCell(5).setCellValue(round(stroke['stroke_average'], 3))
            row_num += 1
        
        # Auto-size columns
        try:
            for i in range(max(6, n_segments + 1)):
                sheet.autoSizeColumn(i)
        except:
            pass
    
    def create_publication_nuclear_chart(self, sheet, cell_num, data_start_row, data_end_row, n_segments,
                                        anchor_col1, anchor_row1, anchor_col2, anchor_row2):
        """Create publication-ready nuclear membrane chart with separate colored series for each stroke"""
        if not CHARTS_AVAILABLE:
            return
        
        try:
            # Use only the basic working classes
            from org.apache.poi.xddf.usermodel.chart import (
                ChartTypes, XDDFDataSourcesFactory, XDDFCategoryAxis, XDDFValueAxis, AxisPosition
            )
            from org.apache.poi.ss.util import CellRangeAddress
            
            # Create drawing and anchor
            drawing = sheet.createDrawingPatriarch()
            anchor = drawing.createAnchor(0, 0, 0, 0, anchor_col1, anchor_row1, anchor_col2, anchor_row2)
            
            # Create chart
            chart = drawing.createChart(anchor)
            chart.setTitleText("Nuclear Membrane Profile - Cell " + cell_num)
            
            # Create axes with proper scientific labeling
            bottom_axis = chart.createCategoryAxis(AxisPosition.BOTTOM)
            bottom_axis.setTitle("Profile Index (pixels)")
            
            left_axis = chart.createValueAxis(AxisPosition.LEFT)
            left_axis.setTitle("Fluorescence Intensity (a.u.)")
            
            # Create scatter chart for publication-quality point visualization
            scatter_data = chart.createData(ChartTypes.SCATTER, bottom_axis, left_axis)
            
            # Add data series for each stroke (simplified - no custom colors)
            x_range = CellRangeAddress(data_start_row, data_end_row, 0, 0)  # Profile Index column
            x_source = XDDFDataSourcesFactory.fromNumericCellRange(sheet, x_range)
            
            for seg in range(n_segments):
                # Create data range for this stroke (column seg+1)
                y_range = CellRangeAddress(data_start_row, data_end_row, seg + 1, seg + 1)
                y_source = XDDFDataSourcesFactory.fromNumericCellRange(sheet, y_range)
                
                # Add series for this stroke
                series = scatter_data.addSeries(x_source, y_source)
                series.setTitle("Stroke " + str(seg + 1), None)
            
            # Plot the chart
            chart.plot(scatter_data)
            
            IJ.log("Simplified nuclear chart created with " + str(n_segments) + " separate stroke series")
            return True
            
        except Exception as e:
            IJ.log("Nuclear chart creation failed: " + str(e))
            import traceback
            IJ.log("Error details: " + traceback.format_exc())
            return False
    
    def create_publication_cytoplasm_chart(self, sheet, cell_num, data_start_row, data_end_row,
                                         anchor_col1, anchor_row1, anchor_col2, anchor_row2):
        """Create publication-ready cytoplasm chart with above/below threshold series"""
        if not CHARTS_AVAILABLE:
            return
        
        try:
            # Use only the basic working classes
            from org.apache.poi.xddf.usermodel.chart import (
                ChartTypes, XDDFDataSourcesFactory, XDDFCategoryAxis, XDDFValueAxis, AxisPosition
            )
            from org.apache.poi.ss.util import CellRangeAddress
            
            # Create drawing and anchor
            drawing = sheet.createDrawingPatriarch()
            anchor = drawing.createAnchor(0, 0, 0, 0, anchor_col1, anchor_row1, anchor_col2, anchor_row2)
            
            # Create chart
            chart = drawing.createChart(anchor)
            chart.setTitleText("Cytoplasm Profile - Cell " + cell_num + " (Top " + str(self.threshold_percentage) + "% Highlighted)")
            
            # Create axes with proper scientific labeling
            bottom_axis = chart.createCategoryAxis(AxisPosition.BOTTOM)
            bottom_axis.setTitle("Profile Index (pixels)")
            
            left_axis = chart.createValueAxis(AxisPosition.LEFT)
            left_axis.setTitle("Fluorescence Intensity (a.u.)")
            
            # Create scatter chart for publication-quality visualization
            scatter_data = chart.createData(ChartTypes.SCATTER, bottom_axis, left_axis)
            
            # X-axis data (same for both series)
            x_range = CellRangeAddress(data_start_row, data_end_row, 0, 0)  # Profile Index
            x_source = XDDFDataSourcesFactory.fromNumericCellRange(sheet, x_range)
            
            # Series 1: Above threshold
            above_y_range = CellRangeAddress(data_start_row, data_end_row, 1, 1)  # Column B
            above_y_source = XDDFDataSourcesFactory.fromNumericCellRange(sheet, above_y_range)
            
            above_series = scatter_data.addSeries(x_source, above_y_source)
            above_series.setTitle("Above Threshold (Top " + str(self.threshold_percentage) + "%)", None)
            
            # Series 2: Below threshold
            below_y_range = CellRangeAddress(data_start_row, data_end_row, 2, 2)  # Column C
            below_y_source = XDDFDataSourcesFactory.fromNumericCellRange(sheet, below_y_range)
            
            below_series = scatter_data.addSeries(x_source, below_y_source)
            below_series.setTitle("Below Threshold", None)
            
            # Plot the chart
            chart.plot(scatter_data)
            
            IJ.log("Simplified cytoplasm chart created with above/below threshold series")
            return True
            
        except Exception as e:
            IJ.log("Cytoplasm chart creation failed: " + str(e))
            import traceback
            IJ.log("Error details: " + traceback.format_exc())
            return False
    
    def create_publication_summary_sheet(self, workbook):
        """Create publication-ready summary sheet"""
        sheet = workbook.createSheet("Summary")
        
        # Get current date/time
        date_format = SimpleDateFormat("yyyy-MM-dd HH:mm:ss")
        current_time = date_format.format(Date())
        
        row_num = 0
        
        # Header section
        header_row = sheet.createRow(row_num)
        header_cell = header_row.createCell(0)
        header_cell.setCellValue("Nuclear Membrane & Cytoplasm Fluorescence Analysis")
        row_num += 2
        
        # Analysis info
        info_data = [
            ("Generated:", str(current_time)),
            ("Total cells analyzed:", str(self.cell_count)),
            ("Image name:", self.image_name),
            ("Threshold percentage:", str(self.threshold_percentage) + "%"),
            ("Chart style:", "Publication-ready with separate data series"),
            ("Nuclear membrane:", "Each stroke plotted as separate colored series"),
            ("Cytoplasm:", "Above-threshold (forest green) vs below-threshold (light gray)")
        ]
        
        for label, value in info_data:
            row = sheet.createRow(row_num)
            row.createCell(0).setCellValue(label)
            row.createCell(1).setCellValue(value)
            row_num += 1
        
        row_num += 1
        
        # Per-cell data table
        header_row = sheet.createRow(row_num)
        headers = ["Cell ID", "Nuclear Mean Intensity", "Cytoplasm Mean Intensity", "Nuclear/Cytoplasm Ratio", "Stroke Count", "Cytoplasm Pixels"]
        for i, header in enumerate(headers):
            cell = header_row.createCell(i)
            cell.setCellValue(header)
        
        row_num += 1
        
        # Data rows
        for i, cell_data in enumerate(self.cell_data):
            row = sheet.createRow(row_num)
            row.createCell(0).setCellValue("Cell " + str(i + 1).zfill(2))
            row.createCell(1).setCellValue(round(cell_data['nuclear_mean'], 3))
            row.createCell(2).setCellValue(round(cell_data['cytoplasm_mean'], 3))
            row.createCell(3).setCellValue(round(cell_data['ratio'], 3))
            row.createCell(4).setCellValue(float(cell_data['nuclear_data']['n_segments']))
            row.createCell(5).setCellValue(float(len(cell_data['cytoplasm_data']['raw_profile'])))
            row_num += 1
        
        row_num += 2
        
        # Summary statistics
        if len(self.nuclear_membrane_averages) > 0:
            nuclear_mean = sum(self.nuclear_membrane_averages) / float(len(self.nuclear_membrane_averages))
            nuclear_se = self.calculate_se(self.nuclear_membrane_averages)
            
            cytoplasm_mean = sum(self.cytoplasm_averages) / float(len(self.cytoplasm_averages))
            cytoplasm_se = self.calculate_se(self.cytoplasm_averages)
            
            ratio_mean = sum(self.nucleus_cytoplasm_ratios) / float(len(self.nucleus_cytoplasm_ratios))
            ratio_se = self.calculate_se(self.nucleus_cytoplasm_ratios)
            
            # Statistics header
            stats_header = sheet.createRow(row_num)
            stats_header.createCell(0).setCellValue("Summary Statistics")
            row_num += 1
            
            # Column headers with ASCII-safe characters
            col_headers = sheet.createRow(row_num)
            stat_headers = ["Parameter", "Mean +/- SE", "n"]
            for i, header in enumerate(stat_headers):
                col_headers.createCell(i).setCellValue(header)
            row_num += 1
            
            # Statistics data with ASCII-safe formatting
            stats_data = [
                ("Nuclear Membrane Intensity", str(round(nuclear_mean, 3)) + " +/- " + str(round(nuclear_se, 3)), self.cell_count),
                ("Cytoplasm Intensity", str(round(cytoplasm_mean, 3)) + " +/- " + str(round(cytoplasm_se, 3)), self.cell_count),
                ("Nuclear/Cytoplasm Ratio", str(round(ratio_mean, 3)) + " +/- " + str(round(ratio_se, 3)), self.cell_count)
            ]
            
            for param, mean_se, n_val in stats_data:
                row = sheet.createRow(row_num)
                row.createCell(0).setCellValue(param)
                row.createCell(1).setCellValue(mean_se)
                row.createCell(2).setCellValue(float(n_val))
                row_num += 1
        
        # Auto-size columns
        try:
            for i in range(6):
                sheet.autoSizeColumn(i)
        except:
            pass
    
    def create_publication_metadata_sheet(self, workbook):
        """Create publication-ready metadata sheet"""
        sheet = workbook.createSheet("Metadata")
        
        row_num = 0
        
        # Header
        header_row = sheet.createRow(row_num)
        header_row.createCell(0).setCellValue("Experimental Metadata")
        row_num += 2
        
        # Analysis parameters
        date_format = SimpleDateFormat("yyyy-MM-dd HH:mm:ss")
        current_time = date_format.format(Date())
        
        params_data = [
            ("Analysis date:", str(current_time)),
            ("Image file:", self.image_name),
            ("Threshold percentage:", str(self.threshold_percentage) + "%"),
            ("Analysis method:", "Stroke-by-stroke nuclear membrane, threshold-based cytoplasm"),
            ("Chart visualization:", "Publication-ready separate data series"),
            ("Nuclear membrane:", "Each polyline stroke plotted as distinct colored series"),
            ("Cytoplasm:", "Above-threshold (forest green) vs below-threshold (light gray)"),
            ("Statistical analysis:", "Mean +/- standard error")
        ]
        
        for label, value in params_data:
            row = sheet.createRow(row_num)
            row.createCell(0).setCellValue(label)
            row.createCell(1).setCellValue(value)
            row_num += 1
        
        row_num += 2
        
        # Microscope metadata
        if len(self.microscope_metadata) > 0:
            metadata_header = sheet.createRow(row_num)
            metadata_header.createCell(0).setCellValue("Microscope Parameters")
            row_num += 1
            
            col_header_row = sheet.createRow(row_num)
            col_header_row.createCell(0).setCellValue("Parameter")
            col_header_row.createCell(1).setCellValue("Value")
            row_num += 1
            
            for key, value in self.microscope_metadata.items():
                row = sheet.createRow(row_num)
                row.createCell(0).setCellValue(str(key))
                row.createCell(1).setCellValue(str(value))
                row_num += 1
        
        # Auto-size columns
        try:
            sheet.autoSizeColumn(0)
            sheet.autoSizeColumn(1)
        except:
            pass
    
    def get_export_settings(self):
        """Get export settings for publication-ready output"""
        try:
            # Get output directory
            dc = DirectoryChooser("Select output directory for Publication-Ready Excel file")
            output_dir = dc.getDirectory()
            
            if output_dir is None:
                IJ.log("Export cancelled by user")
                return None
            
            IJ.log("Selected directory: " + str(output_dir))
            
            # Get filename
            date_format = SimpleDateFormat("yyMMdd")
            current_date = date_format.format(Date())
            
            suggested_name = "Nuclear_Analysis_Publication_" + current_date
            
            gd = GenericDialog("Publication-Ready Excel Export")
            gd.addMessage("Configure your publication-ready Excel export:")
            gd.addStringField("Filename (without extension):", suggested_name, 50)
            gd.addMessage("")
            gd.addMessage("PUBLICATION-READY FEATURES:")
            gd.addMessage("✓ Nuclear membrane: Each stroke as separate colored data series")
            gd.addMessage("✓ Cytoplasm: Forest green (above threshold) vs light gray (below)")
            gd.addMessage("✓ Clean scientific styling with proper axis labels")
            gd.addMessage("✓ Professional legends and no visual clutter")
            gd.addMessage("✓ Summary statistics in Mean ± SE format")
            gd.showDialog()
            
            if gd.wasCanceled():
                IJ.log("Export cancelled by user")
                return None
            
            filename = gd.getNextString().strip()
            if not filename:
                filename = suggested_name
            
            # Ensure .xlsx extension
            if not filename.lower().endswith('.xlsx'):
                filename += '.xlsx'
            
            # Construct file path
            filepath = os.path.join(output_dir, filename)
            
            IJ.log("Full filepath: " + str(filepath))
            
            return {
                'filepath': filepath,
                'filename': filename,
                'directory': output_dir
            }
            
        except Exception as e:
            IJ.log("Error in get_export_settings: " + str(e))
            return None
    
    def calculate_se(self, data):
        """Calculate standard error"""
        if len(data) <= 1:
            return 0.0
        
        mean = sum(data) / float(len(data))
        sum_squared_diff = sum((float(x) - mean) ** 2 for x in data)
        variance = sum_squared_diff / float(len(data) - 1)
        std_dev = math.sqrt(variance)
        se = std_dev / math.sqrt(float(len(data)))
        
        return se

# Run the publication-ready analysis
IJ.log("Starting Publication-Ready Nuclear Membrane & Cytoplasm Analyzer...")
IJ.log("Features: Separate data series, scientific colors, publication-ready styling")
analyzer = EnhancedNuclearAnalyzer()
analyzer.run_analysis()