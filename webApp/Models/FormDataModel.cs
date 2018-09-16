using MultipartDataMediaFormatter.Infrastructure;

namespace webApp.Models
{
    public class FormDataModel
    {
        public string DataJson { get; set; }
        public HttpFile File { get; set; }
    }
}
