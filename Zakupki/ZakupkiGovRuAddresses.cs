namespace Zakupki
{
    internal static class ZakupkiGovRuAddresses
    {
        private const string BASE_URL = "https://zakupki.gov.ru/epz";

        private const string SEARCH_ORDER = "/order/extendedsearch/results.html";
        private const string ORDER_DOCUMENTS = "/order/notice/ea44/view/documents.html";
        private const string ORDER_RESULTS = "/order/notice/rpec/search-results.html";

        private const string REG_NUMBER_PARAM = "regNumber";
        private const string ORDER_NUMBER_PARAM = "orderNum";

        public static readonly string searchAddressPattern = $"{BASE_URL}{SEARCH_ORDER}?{{0}}&recordsPerPage=_500&fz44=on&sortBy=PUBLISH_DATE&publishDateFrom=01.01.2018";
        public static readonly string orderResultAddressPattern = $"{BASE_URL}{ORDER_RESULTS}?{ORDER_NUMBER_PARAM}={{0}}";
        public static readonly string orderDocumentsAddressPattern = $"{BASE_URL}{ORDER_DOCUMENTS}?{REG_NUMBER_PARAM}={{0}}";
    }
}
