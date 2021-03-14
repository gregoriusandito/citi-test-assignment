import React, { useState, useCallback } from 'react';
import clsx from 'clsx';
import PropTypes from 'prop-types';
import {
  Box,
  Card,
  CardContent,
  CardHeader,
  Divider,
  Grid,
  Typography,
  makeStyles
} from '@material-ui/core';
import XLSX from 'xlsx';
import { useDropzone } from 'react-dropzone';
import * as FileSaver from 'file-saver';
import moment from 'moment';

const useStyles = makeStyles(() => ({
  root: {
    height: '100%'
  }
}));

const CustomerData = ({ className, ...rest }) => {
  const classes = useStyles();

  const fileType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8';
  // Desired file extesion
  const fileExtension = '.xlsx';

  const exportToSpreadsheet = (data, rawData, fileName) => {
    // // Create a new Work Sheet using the data stored in an Array of Arrays.
    console.log(data);
    const workSheet = XLSX.utils.aoa_to_sheet(rawData);
    console.log(rawData);
    const workSheet2 = XLSX.utils.aoa_to_sheet(data);
    console.log(workSheet);

    // Generate a Work Book containing the above sheet.
    const workBook = {
      SheetNames: [],
      Sheets: {}
    };

    // const workBook = {};
    XLSX.utils.book_append_sheet(workBook, workSheet, 'Data');
    XLSX.utils.book_append_sheet(workBook, workSheet2, 'Result');

    // Exporting the file with the desired name and extension.
    const excelBuffer = XLSX.write(workBook, { bookType: 'xlsx', type: 'array' });
    const fileData = new Blob([excelBuffer], { type: fileType });
    FileSaver.saveAs(fileData, fileName + fileExtension);
  };

  const fillEmptydata = (data) => {
    const filledData = [];

    data.forEach((value, index) => {
      let tempArray = [];

      if (index !== 0) {
        let name = value[0];
        let dob = value[1];
        let ktp = value[2];
        const date = value[4];

        let [prevName, prevDob, prevKtp] = [];

        if (filledData.length > 0) {
          [prevName, prevDob, prevKtp] = filledData[index - 2];
        }

        if (name === '') {
          name = prevName;
        }

        if (dob === '') {
          dob = prevDob;
        }

        if (ktp === '') {
          ktp = prevKtp;
        }

        tempArray = [name, dob, ktp, date];

        filledData.push(tempArray);
      }
    });

    return filledData;
  };

  // Test
  const onDrop = useCallback((acceptedFiles, fileRejections) => {
    if (fileRejections && fileRejections.length > 0) {
      console.log('Invalid file type. Please upload .xlsx or .xls file.');
    }
    acceptedFiles.forEach((file) => {
      console.log('processing excel file');
      // See https://stackoverflow.com/questions/30859901/parse-xlsx-with-node-and-create-json
      const reader = new FileReader();
      const rABS = !!reader.readAsBinaryString; // converts object to boolean
      reader.onabort = () => {
        console.log('File reading was aborted');
      };
      reader.onerror = () => {
        console.log('file reading has failed');
      };
      reader.onload = (e) => {
        // Do what you want with the file contents
        const bstr = e.target.result;
        const workbook = XLSX.read(bstr, { type: rABS ? 'binary' : 'array' });
        const sheetNameList = workbook.SheetNames[0];
        const jsonFromExcel = XLSX.utils.sheet_to_json(workbook.Sheets[sheetNameList], {
          raw: false,
          dateNF: 'MM-DD-YYYY',
          header: 1,
          defval: ''
        });

        const dirtyData = fillEmptydata(jsonFromExcel);
        const cleanData = dirtyData.filter((value, index) => {
          const [name, dob, ktp] = value;
          return ktp === 'KTP';
        });

        cleanData.sort((a, b) => {
          return b[0] - a[0];
        });

        let tempArray = [];
        const cleanArray = [];

        cleanData.forEach((value, index) => {
          if (tempArray.length === 0) {
            tempArray.push(value);
          } else {
            tempArray.sort((a, b) => {
              return b[0] - a[0];
            });

            if (tempArray.length > 0) {
              const [name, dob, ktp, date] = tempArray[0];
              if (name !== value[0] || index === cleanData.length - 1) {
                tempArray.forEach((val, idx) => {
                  tempArray.sort((a, b) => {
                    return b[3] - a[3];
                  });

                  if (idx < 3) {
                    const [nameVal, dobVal, ktpVal, dateVal] = val;
                    cleanArray.push([[cleanArray.length + 1, nameVal, dobVal, ktpVal].join(' | ')]);
                  }
                });
                tempArray = [];
              }
              tempArray.push(value);
            }
          }
        });

        exportToSpreadsheet(cleanArray, jsonFromExcel, 'Filter_Result_'.concat(moment(new Date()).format('DD_MM_YYYY_HH_mm_ss')));
      };

      if (rABS) reader.readAsBinaryString(file);
      else reader.readAsArrayBuffer(file);
    });
  }, []);

  const { getRootProps, getInputProps } = useDropzone({
    onDrop,
    accept: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet, application/vnd.ms-excel'
  });

  return (
    <Card
      className={clsx(classes.root, className)}
      {...rest}
    >
      <CardHeader title="Upload Excel" />
      <Divider />
      <CardContent>
        <Box
          height={300}
          position="relative"
        >
          <Grid
            item
            md={12}
            xs={12}
          >
            <div
              style={{
                padding: '60px',
                backgroundColor: '#eee',
                height: '100%',
                display: 'flex',
                justifyContent: 'center',
                alignItems: 'center',
                boxShadow: ' 0 0 1px rgb(63 63 68 / 5%), 0 1px 2px 0 rgb(63 63 68 / 15%)',
                borderRadius: '4px'
              }}
              {...getRootProps()}
            >
              <input {...getInputProps()} />
              <Typography
                color="textPrimary"
                variant="h6"
                align="center"
                style={{ cursor: 'pointer', height: '100%' }}
              >
                Drag and drop some files here, or click to select files
              </Typography>
            </div>
          </Grid>
        </Box>
      </CardContent>
    </Card>
  );
};

CustomerData.propTypes = {
  className: PropTypes.string
};

export default CustomerData;
