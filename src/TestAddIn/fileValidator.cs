using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestAddIn
{

    class fileValidator
    {
        public static SolidEdgeDraft.DraftDocument objDraftDocument = null;
        /// <summary>
        /// Funkcja zbierająca podstawowe informacje 
        /// </summary>
        /// <param name="draftDocument">odbiera parametr - aktywny dokument DRAFT</param>
        /// <returns>prawda/fałsz, czy dokument może być prcesowany</returns>
        public static bool getValidationDocumentInfo(SolidEdgeDraft.DraftDocument draftDocument)
        {

            return false;
        }
        /// <summary>
        /// Funkacja walidująca, czy arkusze i operacje na nich są ułożone w kolejności niemalejącej
        /// 
        /// </summary>
        /// <param name="draftDocument"></param>
        /// <returns></returns>
        public static string getStepValidation(SolidEdgeDraft.DraftDocument draftDocument)
        {
            return "";
        }



    }
}
