# Vishal Lall 2014

'''
=============
Instructions
=============
1) Make sure you are in the correct directory and the workbook is NOT open
2) Change 'mmlist' to include all of the smoothing level file names you want to run
3) Change 'subjectList' to include all of the subjects you wants to run
4) Read the Success output. If subject fails, check ROIs and make sure they are included in lines 80-90
To-do: Make Averages worksheet for all runs
'''


import xlsxwriter
import xlrd

mmlist = ['sceneFace_0mm', 'sceneFace_3mm', 'sceneFace_5mm']
subjectList = ['SCM01','SCM02','SCM03','SCM04','SCM05', 'SCM06','SCM07','SCM08','SCM09','SCM10','SCM11', 'SCM12', 'SCM13','SCM14', 'SCM15', 'SCM16']
failure = 0
subjectDict = {}
roiDict = {}
mmDict = {}

for smooth in mmlist:
    wb = xlsxwriter.Workbook('%s_Output.xlsx' % (smooth))

    for subject in subjectList:
        ws = wb.add_worksheet(subject)
        
        path = subject + '/' + smooth +'/'
        with open(path+"NEW_ROI_Names.txt", "r") as myfile:
            roiFile = myfile.readlines()
            myfile.close()
        with open(path+"%s_DYN_func_face.txt" % (subject), "r") as myfile2:
            dynFaceFile = myfile2.readlines()
            myfile2.close()
        with open(path+"%s_DYN_func_scene.txt" % (subject), "r") as myfile3:
            dynSceneFile = myfile3.readlines()
            myfile3.close()
        with open(path+"%s_STAT_func_face.txt" % (subject), "r") as myfile4:
            statFaceFile = myfile4.readlines()
            myfile4.close()
        with open(path+"%s_STAT_func_scene.txt" % (subject), "r") as myfile5:
            statSceneFile = myfile5.readlines()
            myfile5.close()
            
        if (len(roiFile) != len(dynFaceFile)):
            print ('\nWARNING: Your Number of ROIs and data are mismatched!\n')

        # write(Y, X, DATA)
        ws.write(0, 1, 'DYN Scenes')
        ws.write(0, 2, 'STAT Scenes')
        ws.write(0, 3, 'DYN Faces')
        ws.write(0, 4, 'STAT Faces')

        #ROI DATA, Split into lists of regions        
        roiList = []
        lfaceROIList = []
        rfaceROIList = []
        lsceneROIList = []
        rsceneROIList = []

        for i in range (0,len(roiFile)):
            roiData = roiFile[i].rstrip()
            dynFace = dynFaceFile[i].rstrip()
            dynScene = dynSceneFile[i].rstrip()
            statFace = statFaceFile[i].rstrip()
            statScene = statSceneFile[i].rstrip()
            print ( "%s,%s,%s.%s,%s" % (roiData,dynFace,dynScene,statFace,statScene) )
            
            #DICTIONARY: Very useful if we want to compare across excel sheets/subjects
            roiDict[roiData] = [dynFace, dynScene, statFace, statScene]
            subjectDict[subject] = roiDict
            mmDict[smooth] = subjectDict

            #ROI DATA, Split into lists of regions
            # IMPORTANT: Make sure all possible ROI Names are accounted for here or you will recieve a 'FAILURE' Error when running the script, and your data will be mismatched.
            newROI = roiFile[i].rstrip()
            roiList.append(newROI)
            if newROI in ['lFFA','laFFA','lpFFA','lSTS','lpSTS','laSTS','lOFA','laOFA','lpOFA', 'lOFAunparcelated']:
                lfaceROIList.append(newROI)
            elif newROI in ['rFFA','raFFA','rpFFA','rSTS','rpSTS','raSTS','rOFA','raOFA','rpOFA']:
                rfaceROIList.append(newROI)
            elif newROI in ['lOPA', 'lOPAredo','laOPA','lpOPA', 'lOPAunparcelated','lRSC','lpRSC','laRSC','lPPA','laPPA','lpPPA']:
                lsceneROIList.append(newROI)
            elif newROI in ['rOPA','raOPA','rpOPA', 'rOPAlat', 'rOPAmed', 'rOPAredo','rRSC','rpRSC','raRSC', 'rsRSC', 'rlRSC', 'riRSC','rPPA','raPPA','rpPPA']:
                rsceneROIList.append(newROI)

            # (row, col, data)
            # skip one row for the labels
            j = i+1
            ws.write(j, 0, roiData)
            ws.write(j, 1, float(dynScene))
            ws.write(j, 2, float(statScene))
            ws.write(j, 3, float(dynFace))
            ws.write(j, 4, float(statFace)) 

        # There isn't an equal number of ROIS for each area so we offset the ROIs in excel mathematically
        sp = 2
        lf = len(lfaceROIList)
        rf = len(rfaceROIList)
        ls = len(lsceneROIList)
        rs = len(rsceneROIList)
        fp = sp + lf - 1
        np = fp + rf
        np2 = np + ls
        np3 = np2 + rs 

        # FOR LEFT FACES
        chart1 = wb.add_chart({'type': 'column'})
        chart1.add_series({
            'categories': '=%s!$A$%s:$A$%s' % (subject, str(sp), str(fp) ),
            'values': '=%s!$D$%s:$D$%s' % (subject, str(sp), str(fp) ),
            'name': '=%s!$D$1' % (subject),
            'y_error_bars': {'type': 'standard_error'},
        })
        chart1.add_series({
            'categories': '=%s!$A$%s:$A$%s' % (subject, str(sp), str(fp) ),
            'values': '=%s!$E$%s:$E$%s' % (subject, str(sp), str(fp) ),
            'name': '=%s!$E$1' % (subject),
            'y_error_bars': {'type': 'standard_error'},
        })        
        ws.insert_chart('H1', chart1)

        #FOR RIGHT FACES
        chart2 = wb.add_chart({'type': 'column'})
        chart2.add_series({
            'categories': '=%s!$A$%s:$A$%s' % (subject, str(fp+1), str(np) ),
            'values': '=%s!$D$%s:$D$%s' % (subject, str(fp+1), str(np) ),
            'name': '=%s!$D$1' % (subject),
            'y_error_bars': {'type': 'standard_error'},
        })
        chart2.add_series({
            'categories': '=%s!$A$%s:$A$%s' % (subject, str(fp+1), str(np) ),
            'values': '=%s!$E$%s:$E$%s' % (subject, str(fp+1), str(np) ),
            'name': '=%s!$E$1' % (subject),
            'y_error_bars': {'type': 'standard_error'},
        })
        ws.insert_chart('P1', chart2)
        #FOR LEFT SCENES
        chart3 = wb.add_chart({'type': 'column'})
        chart3.add_series({
            'categories': '=%s!$A$%s:$A$%s' % (subject, str(np+1), str(np2) ),
            'values': '=%s!$B$%s:$B$%s' % (subject, str(np+1), str(np2) ),
            'name': '=%s!$B$1' % (subject),
            'y_error_bars': {'type': 'standard_error'},
        })
        chart3.add_series({
            'categories': '=%s!$A$%s:$A$%s' % (subject, str(np+1), str(np2) ),
            'values': '=%s!$C$%s:$C$%s' % (subject, str(np+1), str(np2) ),
            'name': '=%s!$C$1' % (subject),
            'y_error_bars': {'type': 'standard_error'},
        })
        ws.insert_chart('H18', chart3)
        #FOR RIGHT SCENES
        chart4 = wb.add_chart({'type': 'column'})
        chart4.add_series({
            'categories': '=%s!$A$%s:$A$%s' % (subject, str(np2+1), str(np3) ),
            'values': '=%s!$B$%s:$B$%s' % (subject, str(np2+1), str(np3) ),
            'name': '=%s!$B$1' % (subject),
            'y_error_bars': {'type': 'standard_error'},
        })
        chart4.add_series({
            'categories': '=%s!$A$%s:$A$%s' % (subject, str(np2+1), str(np3) ),
            'values': '=%s!$C$%s:$C$%s' % (subject, str(np2+1), str(np3) ),
            'name': '=%s!$C$1' % (subject),
            'y_error_bars': {'type': 'standard_error'},
        })
        ws.insert_chart('P18', chart4)
        
        # CHECK FOR ERRORS
        if len(roiFile)==len(lfaceROIList)+len(rfaceROIList)+len(lsceneROIList)+len(rsceneROIList):
            print ('*** SUCCESS: %s ROI File matched with ROIs Detected! ***\n\n\n' % (subject))
        else:
            print ('*** FAILED TO MATCH ROIs: %s Data may be mismatched! ***\n\n\n' % (subject))
            failure += 1

    wb.close()
print subjectDict
print ('Output: %s failure(s)' % str(failure))
