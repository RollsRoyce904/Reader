namespace Reader
{
    class DataCell
    {
        public int RowNumber { get; set; }
        public string QuestionCell { get; set; }
        public string AnswerCell { get; set; }





        public DataCell()
        {

        }

        public DataCell(int rowNumber, string questionCell, string answerCell)
        {
            RowNumber = rowNumber;
            QuestionCell = questionCell;
            AnswerCell = answerCell;
        }

    }
}
