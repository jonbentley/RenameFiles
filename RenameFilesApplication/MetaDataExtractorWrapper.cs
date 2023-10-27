using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using MetadataExtractor;
using MetadataExtractor.Formats.Exif;

namespace RenameFilesApplication
{
    /// <summary>
    /// The examples for using the Nuget Package meatdata extractor are both brief and in C#
    /// so creating a C# project to save converting
    /// 
    /// https://github.com/drewnoakes/metadata-extractor-dotnet
    /// https://github.com/drewnoakes/metadata-extractor/wiki/Getting-Started-(dotnet)
    /// </summary>
    public class MetaDataExtractorWrapper
    {
        /// <summary>
        /// Get a list of all exif values
        /// </summary>
        /// <param name="imagePath"></param>
        /// <returns></returns>
        public string GetListOfAllExifValues(string imagePath)
        {
            var listOfValues = new StringBuilder();
            listOfValues.AppendLine("_____________________________________________________________________________________________");
            listOfValues.AppendLine("File Path: " + imagePath);

            IEnumerable<Directory> directories = ImageMetadataReader.ReadMetadata(imagePath);
            // The resulting directories sequence holds potentially many different directories of metadata, depending upon the input image.

            foreach (var directory in directories)
            {
                listOfValues.AppendLine("_______________________________");
                listOfValues.AppendLine($"Directory: {directory.Name}");

                foreach (var tag in directory.Tags)
                {
                    listOfValues.AppendLine($"{directory.Name} - {tag.Name} = {tag.Description}");
                }
            }

            var subIfdDirectory = directories.OfType<ExifSubIfdDirectory>().FirstOrDefault();
            var dateTime = subIfdDirectory?.GetDescription(ExifDirectoryBase.TagDateTime);

            return listOfValues.ToString();
        }

        /// <summary>
        /// Get a list of all exif values that are dates
        /// </summary>
        /// <param name="imagePath"></param>
        /// <returns></returns>
        public string GetListOfAllExifDateValues(string imagePath)
        {
            var listOfValues = new StringBuilder();
            listOfValues.AppendLine("_____________________________________________________________________________________________");
            listOfValues.AppendLine("File Path: " + imagePath);
            listOfValues.AppendLine("_______________________________");

            IEnumerable<Directory> directories = ImageMetadataReader.ReadMetadata(imagePath);
            // The resulting directories sequence holds potentially many different directories of metadata, depending upon the input image.

            foreach (var directory in directories)
            {
                foreach (var tag in directory.Tags)
                {
                    if (tag.Name.ToUpper().Contains("DATE") && tag.Description != "0")
                    {
                        listOfValues.AppendLine($"{directory.Name} - {tag.Name} = {tag.Description}");
                    }
                }
            }

            return listOfValues.ToString();
        }


        /// <summary>
        /// Try and get the date time the photo was taken
        /// If we get the date from the 'date photo taken' attribute then all is good and return 'true'
        /// If we couldn't then the situation could be messy/fail so return 'false' to indicate you cannot trust the answer, but do the best we can manage
        /// Return date.min if you cann't find one at all
        /// </summary>
        /// <param name="imagePath"></param>
        /// <param name="datePhotoTaken"></param>
        /// <returns></returns>
        public bool GetExifDateTimePhotoTaken(string imagePath, out DateTime datePhotoTaken)
        {
            IEnumerable<Directory> directories = ImageMetadataReader.ReadMetadata(imagePath);

            // The resulting directories sequence holds potentially many different directories of metadata, depending upon the input image.
            var exifSubIfdDirectory = directories.OfType<ExifSubIfdDirectory>().FirstOrDefault();

            // string dateTimeString = subIfdDirectory?.GetDescription(ExifDirectoryBase.TagDateTime);

            // first try and us the normal field for a photo date
            if (exifSubIfdDirectory != null)
            {
                try
                {
                    exifSubIfdDirectory.TryGetDateTime(ExifDirectoryBase.TagDateTimeOriginal, out DateTime tagDateTimeOriginal);

                    if (tagDateTimeOriginal != DateTime.MinValue)
                    {
                        datePhotoTaken = tagDateTimeOriginal;
                        return true;
                    }
                }

                catch { }

                // if the first failed, this is unlikely to work, try the digitized date
                try
                {
                    exifSubIfdDirectory.TryGetDateTime(ExifDirectoryBase.TagDateTimeDigitized, out DateTime tagDateTimeDigitized);

                    if (tagDateTimeDigitized != DateTime.MinValue)
                    {
                        datePhotoTaken = tagDateTimeDigitized;
                        return false;
                    }
                }

                catch { }

                // last change at a 'named field' the file date
                try
                {
                    var exifIfd0Directory = directories.OfType<ExifIfd0Directory>().FirstOrDefault();
                    exifIfd0Directory.TryGetDateTime(ExifDirectoryBase.TagDateTime, out var tagDateTime);

                    if (tagDateTime != DateTime.MinValue)
                    {
                        datePhotoTaken = tagDateTime;
                        return false;
                    }
                }

                catch { }
            }


            // having failed to get to any named date field, now lets just try any date field
            // loop round all fields looking for a date

            List<string> possibleDateStrings = new List<string>();

            foreach (var directory in directories)
            {
                foreach (var tag in directory.Tags)
                {
                    string dateString = tag.Description;
                    if (tag.Name.ToUpper().Contains("DATE") && tag.Description != "0")
                    {
                        possibleDateStrings.Add(tag.Description);
                    }
                }
            }


            // We have a list of strings that MAY be a useful date. Try them one by one

            foreach (string possibleDate in possibleDateStrings)
            {

                DateTime mydate = new DateTime();

                // try a basic date
                if (DateTime.TryParse(possibleDate, out mydate) == true)
                {
                    datePhotoTaken = mydate;
                    return false;
                }

                // it's not an 'easy date format so lets try a custom date format
                // try a date of format:   2020:12:31 14:12:15
                IFormatProvider provider = System.Globalization.CultureInfo.InvariantCulture.DateTimeFormat;
                if (DateTime.TryParseExact(possibleDate, "yyyy:MM:dd HH:mm:ss", provider, System.Globalization.DateTimeStyles.AllowWhiteSpaces, out mydate) == true)
                {
                    datePhotoTaken = mydate;
                    return false;
                }

                // need to handle - File - File Modified Date = Sat Oct 10 17:24:39 +01:00 2020
                string dateString = possibleDate;
                dateString = dateString.Replace("-01:00", "");
                dateString = dateString.Replace("+00:00", "");
                dateString = dateString.Replace("+01:00", "");
                dateString = dateString.Replace("+02:00", "");


                if (DateTime.TryParseExact(dateString, "ddd MMM dd HH:mm:ss yyyy", provider, System.Globalization.DateTimeStyles.AllowWhiteSpaces, out mydate) == true)
                {
                    datePhotoTaken = mydate;
                    return false;
                }
            }

            // all have failed, return min value
            datePhotoTaken = DateTime.MinValue;
            return false;
        }
    }
}
