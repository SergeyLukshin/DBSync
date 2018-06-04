using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using V82;
using System.Runtime.InteropServices;
using System.Threading;

namespace DBSync
{
    public partial class DBSync : Form
    {
        public DBSync()
        {
            InitializeComponent();
        }

        class ShopInfo
        {
            public string strCode;
            public string strPath;
            public DateTime dateSync;
            //public DateTime dateSyncDocuments;
        }

        // ------------------
        // Номенклатуры
        // ------------------

        class TypeInfo
        {
            public string strCodeSync;
            public string strName;
            public bool bInvisible;
        }

        class SizeInfo
        {
            public string strSize;
            public int iRussianSize;
        }

        class BrandInfo
        {
            //public dynamic ref_;
            public string strCodeSync;
            public string strName;
            public int iSortPos;
            public DateTime dateChange;
            public bool bDeleted;
            public List<SizeInfo> listSizes = new List<SizeInfo>();
        }

        class CollectionInfo
        {
            public string strCodeSync;
            public string strName;
            public int iSortPos;
            public bool bInvisible;
        }

        class GroupInfo
        {
            public string strCodeSync;
            public string strName;
            public BrandInfo brand = new BrandInfo();
            //public string strBrandCodeSync;
            public CollectionInfo collection = new CollectionInfo();
            public CollectionInfo collectionForSite = new CollectionInfo();
            public double fDiscount;
            public double fDiscountWholesale;
            public bool bInvisible;
            public bool bOnlyForRegistered;
            public bool bNewCollection;
        }

        class AttribInfo
        {
            public string strCodeSync;
            public string strName1;
            public string strName2;
            public bool bDeleted;
            public int iRussianSize;
        }

        class TovarInfo
        {
            public string strCodeSync;
            public DateTime dateChange;
            public DateTime dateChangeImages;
            public DateTime dateChangeAttribs;
            public string strArticul;
            public string strName;
            public string strFullName;
            public bool bDeleted;
            public GroupInfo group = new GroupInfo();
            public TypeInfo type = new TypeInfo();
            public dynamic ref_;
            public List<AttribInfo> listAttribs = new List<AttribInfo>();
            public List<AttribInfo> listAttribsShop = new List<AttribInfo>();

            public string strNote;
            public string strNameForSite;
            public string strNoteForSite;

            public double fDiscount;
            public bool bBigSize;
            public bool bNew;
            public double fPriceInEuro;
            public bool bInvisible;
            public int iSort;
        }

        // ------------------
        // Акции и скидки
        // ------------------
        class DiscountInfo
        {
            public string strName;
            public double fProcent;
        }

        class ActionInfo
        {
            public string strName;
            public DateTime dateBegin;
            public DateTime dateEnd;
            public bool bProvodka;
            public bool bDeleted;
            public dynamic ref_ = null;

            public List<DiscountInfo> listDiscounts = new List<DiscountInfo>();
        }

        // ------------------
        // Документы установки цен
        // ------------------

        class TovarPriceInfo
        {
            public string strTovarSyncCode;
            public string strAttribSyncCode;
            public double price;
        }

        class PriceDocumentInfo
        {
            public string strCodeSync;
            public string strCode;
            public DateTime date;
            //public DateTime dateSync;
            public string strNote;
            public bool bProvodka;
            public bool bDeleted;
            public List<TovarPriceInfo> listTovars = new List<TovarPriceInfo>();
        }

        // ------------------
        // Товары в документах
        // ------------------

        class TovarInvoiceInfo
        {
            public string strTovarSyncCode;
            public string strAttribSyncCode;
            public int iCount;
            public double price;
            public double manual_discount;
            public double manual_discount_opt;
            public double summa_manual_discount;
            //public string strInvoiceInCodeSync; // для возврата от покупателя
        }

        // ------------------
        // Контрагенты
        // ------------------
        class SupplierInfo
        {
            public string strCodeSync;
            public string strName;
            public DateTime dateChange;
            public int Type;
            public string strFullName;
            public string strINN;
            public string strNote;

            public string strSurname;
            public string strFirstname;
            public string strFathername;

            public bool bWholeSaler;
            public bool bIgnor;
            public bool bNoAgreeWithDelivery;

            public string strBrandWishes;
            public string strSizeWishes;
            public string strCategoryWishes;

            public string strWishes1;
            public string strWishes2;
            public string strWishes3;

            public Dictionary<string, string> dictExInfo = new Dictionary<string, string>();
        }

        // ------------------
        // Документы прихода
        // ------------------

        class InvoiceInDocumentInfo
        {
            public string strCodeSync;
            public string strCode;
            public DateTime date;
            public DateTime dateChange;
            public bool bProvodka;
            public bool bDeleted;
            public bool bReport;
            public string strSupplierCodeSync;
            public List<TovarInvoiceInfo> listTovars = new List<TovarInvoiceInfo>();
            public string strNote;
            public string strUser;
        }

        // ------------------
        // Документы реализации
        // ------------------

        class DiscountCardInfo
        {
            public string strName;
            public string strCode;
            //public double fProcent;
        }

        class InvoiceOutDocumentInfo
        {
            //public dynamic ref_ = null;
            public string strCodeSync;
            public string strCode;
            public DateTime date;
            public DateTime dateChange;
            public bool bProvodka;
            public bool bDeleted;
            public bool bReport;
            public bool bOptovik;
            public double fDocumentSum; // для контроля скидок
            public string strSupplierCodeSync;
            public DiscountCardInfo card = new DiscountCardInfo();
            public List<TovarInvoiceInfo> listTovars = new List<TovarInvoiceInfo>();
            public string strNote;
            public string strDogovor;
            public string strUser;
        }

        class InvoiceReturnDocumentInfo
        {
            //public dynamic ref_ = null;
            public string strCodeSync;
            public string strCode;
            public DateTime date;
            public DateTime dateChange;
            public bool bProvodka;
            public bool bDeleted;
            public bool bReport;
            public bool bOptovik;
            public bool bManualCost;
            public string strAnaliticName;
            public string strSupplierCodeSync;
            public DiscountCardInfo card = new DiscountCardInfo();
            public List<TovarInvoiceInfo> listTovars = new List<TovarInvoiceInfo>();
            public string strNote;
            public string strUser;
        }

        class InvoiceMoveInOutDocumentInfo
        {
            //public dynamic ref_ = null;
            public string strCodeSync;
            public string strCode;
            public DateTime date;
            public DateTime dateChange;
            public bool bProvodka;
            public bool bDeleted;
            public List<TovarInvoiceInfo> listTovars = new List<TovarInvoiceInfo>();
            public string strNote;
            public string strUser;
        }

        /*class PairTovarAttribInfo
        {
            public PairTovarAttribInfo(int TovarCode, int AttribCode)
            {
                iTovarCode = TovarCode;
                iAttribCode = AttribCode;
            }

            public int iTovarCode;
            public int iAttribCode;
        }*/

        Dictionary<string, ShopInfo> dictShops = new Dictionary<string, ShopInfo>();
        List<TovarInfo> listTovars = new List<TovarInfo>();
        List<TovarInfo> listTovarsShop = new List<TovarInfo>();
        
        List<BrandInfo> listBrands = new List<BrandInfo>();
        List<BrandInfo> listBrandsShop = new List<BrandInfo>();

        List<PriceDocumentInfo> listPriceDocuments = new List<PriceDocumentInfo>();
        List<PriceDocumentInfo> listPriceDocumentsShop = new List<PriceDocumentInfo>();

        List<SupplierInfo> listSuppliers = new List<SupplierInfo>();
        List<SupplierInfo> listSuppliersShop = new List<SupplierInfo>();

        //List<InvoiceInDocumentInfo> listInvoiceInDocuments = new List<InvoiceInDocumentInfo>();
        List<InvoiceInDocumentInfo> listInvoiceInDocumentsShop = new List<InvoiceInDocumentInfo>();

        List<InvoiceOutDocumentInfo> listInvoiceOutDocumentsShop = new List<InvoiceOutDocumentInfo>();

        List<InvoiceReturnDocumentInfo> listInvoiceReturnDocumentsShop = new List<InvoiceReturnDocumentInfo>();

        List<InvoiceMoveInOutDocumentInfo> listInvoiceMoveInDocumentsShop = new List<InvoiceMoveInOutDocumentInfo>();
        List<InvoiceMoveInOutDocumentInfo> listInvoiceMoveOutDocumentsShop = new List<InvoiceMoveInOutDocumentInfo>();

        Dictionary<string, int> dictPairTovarAttribRemain = new Dictionary<string, int>();

        //public bool bTest = true;

        class PairTovarAttribInfo
        {
            public PairTovarAttribInfo()
            {
            }
            public int iRemain;
            public double fCost;
            public string strTovarName;
            public string strAttribName;
            public bool bTovarDeleted;
            public bool bAttribDeleted;
            public string iTovarCode;
            public string iAttribCode;
        }
        Dictionary<string, PairTovarAttribInfo> dictPairTovarAttribInfo = new Dictionary<string, PairTovarAttribInfo>();
        Dictionary<string, PairTovarAttribInfo> dictPairTovarAttribInfoShop = new Dictionary<string, PairTovarAttribInfo>();


        class VerifyInvoiceInfo
        {
            public enum InvoiceType
            {
                InvoiceIn = 1,
                InvoiceOut = 2,
                Return = 3,
                MoveIn = 4,
                MoveOut = 5
            }

            public string strCodeSync;
            public string strCode;
            public InvoiceType type;
            public DateTime date;
            public bool bProvodka;
            public bool bDeleted;
            public double fCost;
        }
        List<VerifyInvoiceInfo> listVerifyInvoiceShop = new List<VerifyInvoiceInfo>();

        ActionInfo ActionDocument = null;
        ActionInfo ActionShopDocument = null;

        V82.COMConnector v8con;// = new V82.COMConnector();
        dynamic connect;

        V82.COMConnector v8conShop;// = new V82.COMConnector();
        dynamic connectShop;

        dynamic ref_shop = null;
        dynamic ref_company = null;
        //dynamic ref_warehouse = null;

        dynamic findDiscount;
        dynamic findDiscountCond;
        dynamic findDiscountCard;

        dynamic brand;
        dynamic tovar;
        dynamic attrib;

        string m_strSelectedShop;
        bool m_bDeleteAutoInvoice;
        bool m_bSaveInProgress = false;
        bool m_bVerifyMainBase = false;
        bool m_bNoDeleteDocuments = false;
        bool m_bWasShopSelect = false;

        DateTime DateSync;

        private bool EqualActionDiscounts()
        {
            for (int i = 0; i < ActionDocument.listDiscounts.Count; i++)
            {
                bool bFind = false;
                for (int j = 0; j < ActionShopDocument.listDiscounts.Count; j++)
                {
                    if (ActionDocument.listDiscounts[i].strName == ActionShopDocument.listDiscounts[j].strName &&
                        ActionDocument.listDiscounts[i].fProcent == ActionShopDocument.listDiscounts[j].fProcent)
                    {
                        bFind = true;
                        break;
                    }
                }

                if (!bFind) return false;
            }

            return true;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (cbShop.SelectedIndex < 0)
            {
                MessageBox.Show("Необходимо выбрать магазин.");
                return;
            }

            if (connectShop != null) Marshal.FinalReleaseComObject(connectShop);
            if (v8conShop != null) Marshal.FinalReleaseComObject(v8conShop);

            v8conShop = new V82.COMConnector();
            v8conShop.PoolCapacity = 10;
            v8conShop.PoolTimeout = 60;
            string strConnectionString = tbShopBaseName.Text;
            try
            {
                connectShop = v8conShop.Connect(strConnectionString);
            }
            catch (Exception ex)
            {
                Cursor.Current = Cursors.Default;

                Marshal.FinalReleaseComObject(v8conShop);
                MessageBox.Show(ex.Message);
                return;
            }

            listView1.Items.Clear();

            DateSync = dictShops[cbShop.Items[cbShop.SelectedIndex].ToString()].dateSync;

            btnFixRemains.Enabled = false;
            cbDeleteAutoInvoice.Enabled = false;
            button1.Enabled = false;

            cbVerifyMainBase.Enabled = false;
            btnRefreshDataSite.Enabled = false;
            btnYML.Enabled = false;

            m_bWasShopSelect = true;

            backgroundWorker3.RunWorkerAsync();
       
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;

            if (connect != null) Marshal.FinalReleaseComObject(connect);
            if (v8con != null) Marshal.FinalReleaseComObject(v8con);

            v8con = new V82.COMConnector();
            string strConnectionString = tbMainBaseName.Text;
            v8con.PoolCapacity = 10;
            v8con.PoolTimeout = 60; 
            try
            {
                connect = v8con.Connect(strConnectionString);
            }
            catch (Exception ex)
            {
                Cursor.Current = Cursors.Default;

                Marshal.FinalReleaseComObject(v8con);
                MessageBox.Show(ex.Message);
                return;
            }

            try
            {
                gbShop.Enabled = true;
                btnRefreshDataSite.Enabled = true;
                btnYML.Enabled = true;
                cbShop.Items.Clear();
                dynamic shops = connect.Справочники.Магазины.Выбрать();
                while (shops.Следующий())
                {
                    if (!shops.ПометкаУдаления)
                    {
                        string strName = shops.Наименование;
                        cbShop.Items.Add(strName);
                        dictShops[strName] = new ShopInfo();
                        dictShops[strName].strCode = shops.Код;
                        dictShops[strName].strPath = shops.БазаМагазина;
                        dictShops[strName].dateSync = shops.ДатаСинхронизации;
                    }
                }

                if (cbShop.Items.Count == 1)
                {
                    cbShop.SelectedIndex = 0;
                    tbShopBaseName.Text = dictShops[cbShop.Items[0].ToString()].strPath;
                    tbLastDateSync.Text = dictShops[cbShop.Items[0].ToString()].dateSync.ToShortDateString() +
                        " " + dictShops[cbShop.Items[0].ToString()].dateSync.ToShortTimeString();
                }

                Marshal.FinalReleaseComObject(shops);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return;
            }

            Cursor.Current = Cursors.Default;
        }

        private void cbShop_SelectedIndexChanged(object sender, EventArgs e)
        {
            tbShopBaseName.Text = dictShops[cbShop.Items[cbShop.SelectedIndex].ToString()].strPath;
            tbLastDateSync.Text = dictShops[cbShop.Items[cbShop.SelectedIndex].ToString()].dateSync.ToShortDateString() +
                " " + dictShops[cbShop.Items[cbShop.SelectedIndex].ToString()].dateSync.ToShortTimeString();
        }

        private void DBSync_FormClosing(object sender, FormClosingEventArgs e)
        {
            for (int i = 0; i < listTovars.Count(); i++)
            {
                Marshal.FinalReleaseComObject(listTovars[i].ref_);                
            }

            for (int i = 0; i < listTovarsShop.Count(); i++)
            {
                Marshal.FinalReleaseComObject(listTovarsShop[i].ref_);
            }

            listInvoiceOutDocumentsShop.Clear();

            if (ActionDocument != null && ActionDocument.ref_ != null)
            {
                Marshal.FinalReleaseComObject(ActionDocument.ref_);
            }
            if (ActionShopDocument != null && ActionShopDocument.ref_ != null)
            {
                Marshal.FinalReleaseComObject(ActionShopDocument.ref_);
            }

            if (ref_shop != null)
            {
                Marshal.FinalReleaseComObject(ref_shop);
            }
            if (ref_company != null)
            {
                Marshal.FinalReleaseComObject(ref_company);
            }
            if (connectShop != null)
            {
                Marshal.FinalReleaseComObject(connectShop);
            }
            if (v8conShop != null)
            {
                Marshal.FinalReleaseComObject(v8conShop);
            }

            if (connect != null)
            {
                Marshal.FinalReleaseComObject(connect);
            }
            if (v8con != null)
            {
                Marshal.FinalReleaseComObject(v8con);
            }
        }

        private void SavePictureShop(string tmp_dir, int index)
        {
            dynamic pic = connectShop.Справочники.Файлы.СоздатьЭлемент();
            pic.ВладелецФайла = tovar.Ссылка;
            pic.Наименование = tovar.Код;
            pic.ДатаСоздания = DateTime.Now;
            pic.Автор = connectShop.Справочники.Пользователи.НайтиПоНаименованию("Admin", true);
            //pic.ПолноеНаименование = tovar.Код;
            pic.Записать();

            dynamic ver = connectShop.Справочники.ВерсииФайлов.СоздатьЭлемент();
            ver.ТипХраненияФайла = connectShop.Перечисления.ТипыХраненияФайлов.ВИнформационнойБазе;
            ver.Наименование = tovar.Код;
            ver.Владелец = pic.Ссылка;

            dynamic K1 = connectShop.NewObject("Картинка", tmp_dir, true).ПолучитьДвоичныеДанные();
            ver.ФайлХранилище = connectShop.NewObject("ХранилищеЗначения", K1);

            ver.Расширение = "jpg";
            ver.СтатусИзвлеченияТекста = connectShop.Перечисления.СтатусыИзвлеченияТекстаФайлов.НеИзвлечен;
            dynamic file = connectShop.NewObject("Файл", tmp_dir);
            ver.Размер = file.Размер();

            ver.Записать();

            pic.ТекущаяВерсия = ver.Ссылка;
            pic.Записать();

            switch (index)
            {
                case 1:
                    tovar.ФайлКартинки = pic.Ссылка;
                    break;
                case 2:
                    tovar.ФайлКартинки2 = pic.Ссылка;
                    break;
                case 3:
                    tovar.ФайлКартинки3 = pic.Ссылка;
                    break;
                case 4:
                    tovar.ФайлКартинки4 = pic.Ссылка;
                    break;
                case 5:
                    tovar.ФайлКартинки5 = pic.Ссылка;
                    break;
            }

            if (file != null) Marshal.FinalReleaseComObject(file);
            if (K1 != null) Marshal.FinalReleaseComObject(K1);
            if (ver != null) Marshal.FinalReleaseComObject(ver);
            if (pic != null) Marshal.FinalReleaseComObject(pic);
        }

        private void SavePicture(string tmp_dir, int index)
        {
            dynamic pic = connect.Справочники.Файлы.СоздатьЭлемент();
            pic.ВладелецФайла = tovar.Ссылка;
            pic.Наименование = tovar.Код;
            pic.ДатаСоздания = DateTime.Now;
            pic.Автор = connect.Справочники.Пользователи.НайтиПоНаименованию("Admin", true);
            //pic.ПолноеНаименование = tovar.Код;
            pic.Записать();

            dynamic ver = connect.Справочники.ВерсииФайлов.СоздатьЭлемент();
            ver.ТипХраненияФайла = connect.Перечисления.ТипыХраненияФайлов.ВИнформационнойБазе;
            ver.Наименование = tovar.Код;
            ver.Владелец = pic.Ссылка;

            dynamic K1 = connect.NewObject("Картинка", tmp_dir, true).ПолучитьДвоичныеДанные();
            ver.ФайлХранилище = connect.NewObject("ХранилищеЗначения", K1);

            ver.Расширение = "jpg";
            ver.СтатусИзвлеченияТекста = connect.Перечисления.СтатусыИзвлеченияТекстаФайлов.НеИзвлечен;
            dynamic file = connect.NewObject("Файл", tmp_dir);
            ver.Размер = file.Размер();

            ver.Записать();

            pic.ТекущаяВерсия = ver.Ссылка;
            pic.Записать();

            switch (index)
            {
                case 1:
                    tovar.ФайлКартинки = pic.Ссылка;
                    break;
                case 2:
                    tovar.ФайлКартинки2 = pic.Ссылка;
                    break;
                case 3:
                    tovar.ФайлКартинки3 = pic.Ссылка;
                    break;
                case 4:
                    tovar.ФайлКартинки4 = pic.Ссылка;
                    break;
                case 5:
                    tovar.ФайлКартинки5 = pic.Ссылка;
                    break;
            }

            if (file != null) Marshal.FinalReleaseComObject(file);
            if (K1 != null) Marshal.FinalReleaseComObject(K1);
            if (ver != null) Marshal.FinalReleaseComObject(ver);
            if (pic != null) Marshal.FinalReleaseComObject(pic);
        }

        private void ExtractAttribs(string strName, int type)
        {
            dynamic size = null;
            dynamic color = null;
            string[] arr = strName.Split(',');
            for (int i = 0; i < arr.Length; i++)
            {
                if (i == 0) // color
                {
                    string strColor = arr[i].Trim();
                    color = connect.Справочники.ЗначенияСвойствОбъектов.НайтиПоНаименованию(strColor, true);
                    if (color == connect.Справочники.ЗначенияСвойствОбъектов.ПустаяСсылка())
                    {
                        dynamic obj = connect.Справочники.ЗначенияСвойствОбъектов.СоздатьЭлемент();
                        obj.Наименование = strColor;
                        obj.Владелец = connect.ПланыВидовХарактеристик.ДополнительныеРеквизитыИСведения.Цвет.Ссылка;
                        obj.Записать();
                        color = obj.Ссылка;
                        if (obj != null) Marshal.FinalReleaseComObject(obj);
                    }
                }
                if (i == 1) // size
                {
                    // из размера убираем российский размер в скобках
                    string strSize = arr[i].Trim();
                    int pos = strSize.IndexOf("(");
                    if (pos >= 0) strSize = strSize.Substring(0, pos).Trim();

                    size = connect.Справочники.ор_Размеры.НайтиПоНаименованию(strSize, true);
                    if (size == connect.Справочники.ор_Размеры.ПустаяСсылка())
                    {
                        dynamic obj = connect.Справочники.ор_Размеры.СоздатьЭлемент();
                        obj.Наименование = strSize;
                        obj.Записать();
                        size = obj.Ссылка;
                        if (obj != null) Marshal.FinalReleaseComObject(obj);
                    }
                }
            }

            if (type == 0)
            {
                dynamic size_ = attrib.ДополнительныеРеквизиты.Добавить();
                size_.Значение = size;
                size_.Свойство = connect.ПланыВидовХарактеристик.ДополнительныеРеквизитыИСведения.Размер;

                dynamic color_ = attrib.ДополнительныеРеквизиты.Добавить();
                color_.Значение = color;
                color_.Свойство = connect.ПланыВидовХарактеристик.ДополнительныеРеквизитыИСведения.Цвет;

                attrib.Записать();

                if (size != null) Marshal.FinalReleaseComObject(size);
                if (color != null) Marshal.FinalReleaseComObject(color);
                if (size_ != null) Marshal.FinalReleaseComObject(size_);
                if (color_ != null) Marshal.FinalReleaseComObject(color_);
                //if (attrib != null) Marshal.FinalReleaseComObject(attrib);
            }
            else
            {
                dynamic row0 = attrib.ДополнительныеРеквизиты.Получить(0);
                row0.Значение = size;
                row0.Свойство = connect.ПланыВидовХарактеристик.ДополнительныеРеквизитыИСведения.Размер;

                dynamic row1 = attrib.ДополнительныеРеквизиты.Получить(1);
                row1.Значение = color;
                row1.Свойство = connect.ПланыВидовХарактеристик.ДополнительныеРеквизитыИСведения.Цвет;

                attrib.Записать();

                if (row0 != null) Marshal.FinalReleaseComObject(row0);
                if (row1 != null) Marshal.FinalReleaseComObject(row1);
                if (size != null) Marshal.FinalReleaseComObject(size);
                if (color != null) Marshal.FinalReleaseComObject(color);
               // if (attrib != null) Marshal.FinalReleaseComObject(attrib);
            }
        }

        private void ExtractAttribsShop(string strName, int type)
        {
            dynamic size = null;
            dynamic color = null;
            string[] arr = strName.Split(',');
            for (int i = 0; i < arr.Length; i++)
            {
                if (i == 0) // color
                {
                    string strColor = arr[i].Trim();
                    color = connectShop.Справочники.ЗначенияСвойствОбъектов.НайтиПоНаименованию(strColor, true);
                    if (color == connectShop.Справочники.ЗначенияСвойствОбъектов.ПустаяСсылка())
                    {
                        dynamic obj = connectShop.Справочники.ЗначенияСвойствОбъектов.СоздатьЭлемент();
                        obj.Наименование = strColor;
                        obj.Владелец = connectShop.ПланыВидовХарактеристик.ДополнительныеРеквизитыИСведения.Цвет.Ссылка;
                        obj.Записать();
                        color = obj.Ссылка;
                        if (obj != null) Marshal.FinalReleaseComObject(obj);
                    }
                }
                if (i == 1) // size
                {
                    // из размера убираем российский размер в скобках
                    string strSize = arr[i].Trim();
                    int pos = strSize.IndexOf("(");
                    if (pos >= 0) strSize = strSize.Substring(0, pos).Trim();

                    size = connectShop.Справочники.ор_Размеры.НайтиПоНаименованию(strSize, true);
                    if (size == connectShop.Справочники.ор_Размеры.ПустаяСсылка())
                    {
                        dynamic obj = connectShop.Справочники.ор_Размеры.СоздатьЭлемент();
                        obj.Наименование = strSize;
                        obj.Записать();
                        size = obj.Ссылка;
                        if (obj != null) Marshal.FinalReleaseComObject(obj);
                    }
                }
            }

            if (type == 0)
            {
                dynamic size_ = attrib.ДополнительныеРеквизиты.Добавить();
                size_.Значение = size;
                size_.Свойство = connectShop.ПланыВидовХарактеристик.ДополнительныеРеквизитыИСведения.Размер;

                dynamic color_ = attrib.ДополнительныеРеквизиты.Добавить();
                color_.Значение = color;
                color_.Свойство = connectShop.ПланыВидовХарактеристик.ДополнительныеРеквизитыИСведения.Цвет;

                attrib.Записать();

                if (size != null) Marshal.FinalReleaseComObject(size);
                if (color != null) Marshal.FinalReleaseComObject(color);
                if (size_ != null) Marshal.FinalReleaseComObject(size_);
                if (color_ != null) Marshal.FinalReleaseComObject(color_);
                //if (attrib != null) Marshal.FinalReleaseComObject(attrib);
            }
            else
            {
                dynamic row0 = attrib.ДополнительныеРеквизиты.Получить(0);
                row0.Значение = size;
                row0.Свойство = connectShop.ПланыВидовХарактеристик.ДополнительныеРеквизитыИСведения.Размер;

                dynamic row1 = attrib.ДополнительныеРеквизиты.Получить(1);
                row1.Значение = color;
                row1.Свойство = connectShop.ПланыВидовХарактеристик.ДополнительныеРеквизитыИСведения.Цвет;

                attrib.Записать();

                if (row0 != null) Marshal.FinalReleaseComObject(row0);
                if (row1 != null) Marshal.FinalReleaseComObject(row1);
                if (size != null) Marshal.FinalReleaseComObject(size);
                if (color != null) Marshal.FinalReleaseComObject(color);
                // if (attrib != null) Marshal.FinalReleaseComObject(attrib);
            }
        }

        private void SyncRemains()
        {
            Cursor.Current = Cursors.WaitCursor;

            try
            {
                connectShop.Константы.ЭтоОсновнаяБаза.Установить(true);
            }
            catch (Exception ex)
            {
                backgroundWorker2.ReportProgress(1, "Ошибка при синхронизации: " + ex.Message);
                return;
            }


            try
            {
                // рассчитываем кол-во номенклатур на основной базе для выбранного магазина
                backgroundWorker2.ReportProgress(0, "");

                if (m_bVerifyMainBase)
                {
                    backgroundWorker2.ReportProgress(0, "Получение фактического остатка товара из основной базы.");

                    dynamic query = connect.NewObject("Запрос");
                    query.Текст = "ВЫБРАТЬ РАЗРЕШЕННЫЕ " +
                        "ХарактеристикиНоменклатуры.Владелец КАК Номенклатура, " +
                        "ХарактеристикиНоменклатуры.Ссылка КАК Характеристика, " +
                        "СУММА(ЕСТЬNULL(Остатки.КоличествоОстаток, 0)) КАК КоличествоОстаток " +
                        "ИЗ " +
                        "Справочник.ХарактеристикиНоменклатуры КАК ХарактеристикиНоменклатуры " +
                        "{ ЛЕВОЕ СОЕДИНЕНИЕ РегистрНакопления.ТоварыНаСкладах.Остатки(, {(Номенклатура).* КАК Номенклатура, (Характеристика).* КАК Характеристика}) КАК Остатки " +
                            "ПО Остатки.Номенклатура = ХарактеристикиНоменклатуры.Владелец И Остатки.Характеристика = ХарактеристикиНоменклатуры.Ссылка} " +
                         "ГДЕ Остатки.Склад.Магазин = &Магазин " +
                        "СГРУППИРОВАТЬ ПО " +
                        "ХарактеристикиНоменклатуры.Владелец, " +
                        "ХарактеристикиНоменклатуры.Ссылка";

                    query.УстановитьПараметр("Магазин", ref_shop);
                    dynamic tovars = query.Выполнить().Выбрать();

                    dictPairTovarAttribRemain.Clear();

                    int i = 0;
                    int cnt = tovars.Количество();
                    while (tovars.Следующий())
                    {
                        i++;
                        //if (connect.ТипЗнч(tovars.Номенклатура) == connect.Тип("СправочникСсылка.Номенклатура"))
                        {
                            string iTovarCode = tovars.Номенклатура.КодДляСинхронизации;
                            string iAttribCode = tovars.Характеристика.КодДляСинхронизации;
                            int iRemain = tovars.КоличествоОстаток;

                            dictPairTovarAttribRemain[iTovarCode + " " + iAttribCode] = iRemain;

                            backgroundWorker2.ReportProgress(-1, "Получение фактического остатка товара из основной базы (" + i.ToString() + " из " + cnt.ToString() + ").");

                        }
                    }

                    //backgroundWorker2.ReportProgress(-1, "Получение фактического остатка товара из основной базы (ok).");

                    if (tovars != null) Marshal.FinalReleaseComObject(tovars);
                    if (query != null) Marshal.FinalReleaseComObject(query);

                    // -------------------------
                    backgroundWorker2.ReportProgress(0, "Сохранение фактического остатка товара в базе магазина.");

                    dynamic queryShop = connectShop.NewObject("Запрос");
                    queryShop.Текст = "ВЫБРАТЬ РАЗРЕШЕННЫЕ " +
                        "ХарактеристикиНоменклатуры.Владелец КАК Номенклатура, " +
                        "ХарактеристикиНоменклатуры.Ссылка КАК Характеристика, " +
                        "СУММА(ХарактеристикиНоменклатуры.Колво) КАК Колво " +
                        "ИЗ " +
                        "Справочник.ХарактеристикиНоменклатуры КАК ХарактеристикиНоменклатуры " +
                        "СГРУППИРОВАТЬ ПО " +
                        "ХарактеристикиНоменклатуры.Владелец, " +
                        "ХарактеристикиНоменклатуры.Ссылка";

                    dynamic tovarsShop = queryShop.Выполнить().Выбрать();

                    i = 0;
                    cnt = tovarsShop.Количество();
                    while (tovarsShop.Следующий())
                    {
                        i++;
                        //if (connectShop.ТипЗнч(tovarsShop.Номенклатура) == connectShop.Тип("СправочникСсылка.Номенклатура"))
                        {
                            string iTovarCode = tovarsShop.Номенклатура.КодДляСинхронизации;
                            string iAttribCode = tovarsShop.Характеристика.КодДляСинхронизации;
                            int iRemain = tovarsShop.Колво;
                            int iFactRemain = 0;

                            if (!dictPairTovarAttribRemain.TryGetValue(iTovarCode + " " + iAttribCode, out iFactRemain))
                            {
                                iFactRemain = 0;
                            }

                            if (iRemain != iFactRemain)
                            {
                                dynamic attrib_ = tovarsShop.Характеристика.ПолучитьОбъект();
                                attrib_.Колво = iFactRemain;
                                attrib_.Записать();
                                if (attrib_ != null) Marshal.FinalReleaseComObject(attrib_);
                            }
                        }

                        backgroundWorker2.ReportProgress(-1, "Сохранение фактического остатка товара в базе магазина (" + i.ToString() + " из " + cnt.ToString() + ").");
                    }

                    //backgroundWorker2.ReportProgress(-1, "Сохранение фактического остатка товара в базе магазина (ok).");

                    if (tovarsShop != null) Marshal.FinalReleaseComObject(tovarsShop);
                    if (queryShop != null) Marshal.FinalReleaseComObject(queryShop);
                }

                // запускаем процедуру удаления всех накладных, не имеющих флага Отчет = истина и создания фиктивных накладных прихода
                // -------------------------
                backgroundWorker2.ReportProgress(0, "Запуск процедуры фиксации остатков.");

                if (!m_bNoDeleteDocuments)
                {
                    dynamic mod = connectShop.SyncDB;
                    int res = mod.ЗафиксироватьОстатки(m_bDeleteAutoInvoice, m_bVerifyMainBase, m_bDeleteAutoInvoice);

                    if (res == 0)
                    {
                        backgroundWorker2.ReportProgress(2, "Фиксация остатков проведена успешно.");
                    }
                    else
                    {
                        if (res == 5)
                            throw new Exception("не найден контрагент для документов прихода");

                        if (res == 6)
                            throw new Exception("не найден контрагент для документов расхода");

                        throw new Exception("ошибка при вызове функции ЗафиксироватьОстатки, код ошибки: " + res.ToString());
                    }

                    if (mod != null) Marshal.FinalReleaseComObject(mod);
                }
                else
                {
                    dynamic mod = connectShop.SyncDB;
                    int res = mod.ЗафиксироватьОстаткиБезУдаления(m_bDeleteAutoInvoice);

                    if (res == 0)
                    {
                        backgroundWorker2.ReportProgress(2, "Фиксация остатков проведена успешно.");
                    }
                    if (mod != null) Marshal.FinalReleaseComObject(mod);
                }
            }
            catch (Exception ex)
            {
                backgroundWorker2.ReportProgress(1, "Ошибка при синхронизации: " + ex.Message);
                Cursor.Current = Cursors.Default;
                return;
            }

            try
            {
                connectShop.Константы.ЭтоОсновнаяБаза.Установить(false);
            }
            catch (Exception ex)
            {
                backgroundWorker2.ReportProgress(1, "Ошибка при синхронизации: " + ex.Message);
                return;
            }

            Cursor.Current = Cursors.Default;
        }

        private void SyncAll()
        {
            Cursor.Current = Cursors.WaitCursor;

            try
            {
                connectShop.Константы.ЭтоОсновнаяБаза.Установить(true);
            }
            catch (Exception ex)
            {
                backgroundWorker1.ReportProgress(1, "Ошибка при синхронизации: " + ex.Message);
                return;
            }

            if (!SyncBrands())
            {
                Cursor.Current = Cursors.Default;
                return;
            };
            if (!SyncTovars())
            {
                Cursor.Current = Cursors.Default;
                return;
            };

            //if (!bTest)
            {
                if (!SyncPrice())
                {
                    Cursor.Current = Cursors.Default;
                    return;
                };
                if (!SyncDiscount())
                {
                    Cursor.Current = Cursors.Default;
                    return;
                };

                if (!SyncContragent())
                {
                    Cursor.Current = Cursors.Default;
                    return;
                };

                if (!SyncInvoiceIn())
                {
                    Cursor.Current = Cursors.Default;
                    return;
                };
                if (!SyncInvoiceOut())
                {
                    Cursor.Current = Cursors.Default;
                    return;
                };
                if (!SyncInvoiceReturn())
                {
                    Cursor.Current = Cursors.Default;
                    return;
                };
                if (!SyncInvoiceMoveIn())
                {
                    Cursor.Current = Cursors.Default;
                    return;
                };
                if (!SyncInvoiceMoveOut())
                {
                    Cursor.Current = Cursors.Default;
                    return;
                };

                backgroundWorker1.ReportProgress(0, "");

                if (!UpdateDateSync())
                {
                    Cursor.Current = Cursors.Default;
                    return;
                };
                if (!UpdateDateSyncShop())
                {
                    Cursor.Current = Cursors.Default;
                    return;
                };
            }

            try
            {
                connectShop.Константы.ЭтоОсновнаяБаза.Установить(false);
            }
            catch (Exception ex)
            {
                backgroundWorker1.ReportProgress(1, "Ошибка при синхронизации: " + ex.Message);
                return;
            }


            Cursor.Current = Cursors.Default;
        }

        private bool UpdateDateSync()
        {
            try
            {
                // записываем дату и время синхронизации
                string strShopCode = dictShops[m_strSelectedShop].strCode;
                dynamic find_shop = connect.Справочники.Магазины.НайтиПоКоду(strShopCode);
                if (find_shop != connect.Справочники.Магазины.ПустаяСсылка())
                {
                    dynamic shop = find_shop.ПолучитьОбъект();
                    shop.ДатаСинхронизации = DateTime.Now;
                    shop.Записать();

                    if (shop != null) Marshal.FinalReleaseComObject(shop);
                }
                if (find_shop != null) Marshal.FinalReleaseComObject(find_shop);

                backgroundWorker1.ReportProgress(2, "Дата синхронизации в основной базе успешно изменена.");

            }
            catch (Exception ex)
            {
                backgroundWorker1.ReportProgress(1, "Ошибка при синхронизации: " + ex.Message);
                //listView1.Items.Add("Ошибка при синхронизации номенклатур и цен: " + ex.Message);
                Cursor.Current = Cursors.Default;
                //MessageBox.Show(ex.Message);
                return false;
            }

            return true;
        }

        private bool UpdateDateSyncShop()
        {
            try
            {
                // записываем дату и время синхронизации
                connectShop.Константы.ДатаСинхронизации.Установить(DateTime.Now);
                backgroundWorker1.ReportProgress(2, "Дата синхронизации в базе магазина успешно изменена.");
            }
            catch (Exception ex)
            {
                backgroundWorker1.ReportProgress(1, "Ошибка при синхронизации: " + ex.Message);
                //listView1.Items.Add("Ошибка при синхронизации номенклатур и цен: " + ex.Message);
                Cursor.Current = Cursors.Default;
                //MessageBox.Show(ex.Message);
                return false;
            }

            return true;
        }

        private bool SyncBrands()
        {
            try
            {
                DateTime date = dictShops[m_strSelectedShop].dateSync;

                if (listBrands.Count() > 0 || listBrandsShop.Count() > 0)
                {
                    backgroundWorker1.ReportProgress(0, "");
                    backgroundWorker1.ReportProgress(0, "Синхронизация номенклатур");

                    bool bExistsCross = false;
                    for (int i = 0; i < listBrands.Count; i++)
                    {
                        for (int j = listBrandsShop.Count - 1; j >= 0; j--)
                        {
                            if (listBrands[i].strCodeSync == listBrandsShop[j].strCodeSync)
                            {
                                //listSuppliersShop.RemoveAt(j);
                                bExistsCross = true;
                            }
                        }
                    }

                    if (bExistsCross)
                    {
                        DialogResult res = MessageBox.Show("Некоторые бренды изменялись и в основной базе и в базе магазина. \nЕсли необходимо брать данные из основной базы - нажмите 'Yes' или 'Да', если из базы магазина - 'No' или 'Нет'", "Предупреждение", MessageBoxButtons.YesNo);
                        if (res == System.Windows.Forms.DialogResult.Yes)
                        {
                            for (int i = 0; i < listBrands.Count; i++)
                            {
                                for (int j = listBrandsShop.Count - 1; j >= 0; j--)
                                {
                                    if (listBrands[i].strCodeSync == listBrandsShop[j].strCodeSync)
                                    {
                                        listBrandsShop.RemoveAt(j);
                                    }
                                }
                            }
                        }
                        else
                        {
                            for (int i = 0; i < listBrandsShop.Count; i++)
                            {
                                for (int j = listBrands.Count - 1; j >= 0; j--)
                                {
                                    if (listBrands[j].strCodeSync == listBrandsShop[i].strCodeSync)
                                    {
                                        listBrands.RemoveAt(j);
                                    }
                                }
                            }
                        }
                    }
                }

                // скидываем данные из основной базы
                for (int i = 0; i < listBrands.Count(); i++)
                {
                    string strCodeSync = listBrands[i].strCodeSync;
                    dynamic find = connectShop.Справочники.ор_ТорговыеМарки.НайтиПоРеквизиту("КодДляСинхронизации", strCodeSync);
                    //dynamic tovar = null;
                    if (find == connectShop.Справочники.ор_ТорговыеМарки.ПустаяСсылка())
                    {
                        // необходимо добавить номенклатуру
                        brand = connectShop.Справочники.ор_ТорговыеМарки.СоздатьЭлемент();
                    }
                    else
                    {
                        // необходимо изменить номенклатуру
                        brand = find.ПолучитьОбъект();
                    }

                    brand.Наименование = listBrands[i].strName;
                    brand.ПометкаУдаления = listBrands[i].bDeleted;
                    brand.Сортировка = listBrands[i].iSortPos;
                    brand.КодДляСинхронизации = listBrands[i].strCodeSync;
                    brand.Синхронизация = true;
                    brand.СоответствиеРазмеров.Очистить();

                    for (int j = 0; j < listBrands[i].listSizes.Count; j++)
                    {
                        // поиск размера
                        // ------------------------
                        dynamic findSize = connectShop.Справочники.ор_Размеры.НайтиПоНаименованию(listBrands[i].listSizes[j].strSize, true);
                        dynamic size = null;
                        if (findSize == connectShop.Справочники.ор_Размеры.ПустаяСсылка())
                        {
                            // необходимо добавить тип
                            size = connectShop.Справочники.ор_Размеры.СоздатьЭлемент();
                            size.Наименование = listBrands[i].listSizes[j].strSize;
                            size.Записать();
                            findSize = size.Ссылка;
                        }
                        // ------------------------

                        dynamic new_row = brand.СоответствиеРазмеров.Добавить();
                        new_row.Размер = findSize;
                        new_row.РусскийРазмер = listBrands[i].listSizes[j].iRussianSize;

                        if (new_row != null) Marshal.FinalReleaseComObject(new_row);
                        if (size != null) Marshal.FinalReleaseComObject(size);
                        if (findSize != null) Marshal.FinalReleaseComObject(findSize);
                    }

                    brand.Записать();

                    if (brand != null) Marshal.FinalReleaseComObject(brand);
                    if (find != null) Marshal.FinalReleaseComObject(find);

                    backgroundWorker1.ReportProgress(0, (i + 1).ToString() + ". Бренд " + listBrands[i].strName + " успешно синхронизирован.");
                }

                // скидываем данные из базы магазина
                for (int i = 0; i < listBrandsShop.Count(); i++)
                {
                    string strCodeSync = listBrandsShop[i].strCodeSync;
                    dynamic find = connect.Справочники.ор_ТорговыеМарки.НайтиПоРеквизиту("КодДляСинхронизации", strCodeSync);
                    //dynamic tovar = null;
                    if (find == connect.Справочники.ор_ТорговыеМарки.ПустаяСсылка())
                    {
                        // необходимо добавить номенклатуру
                        brand = connect.Справочники.ор_ТорговыеМарки.СоздатьЭлемент();
                    }
                    else
                    {
                        // необходимо изменить номенклатуру
                        brand = find.ПолучитьОбъект();
                    }

                    brand.Наименование = listBrandsShop[i].strName;
                    brand.ПометкаУдаления = listBrandsShop[i].bDeleted;
                    brand.Сортировка = listBrandsShop[i].iSortPos;
                    brand.КодДляСинхронизации = listBrandsShop[i].strCodeSync;
                    brand.Синхронизация = true;
                    brand.СоответствиеРазмеров.Очистить();

                    for (int j = 0; j < listBrandsShop[i].listSizes.Count; j++)
                    {
                        // поиск размера
                        // ------------------------
                        dynamic findSize = connect.Справочники.ор_Размеры.НайтиПоНаименованию(listBrandsShop[i].listSizes[j].strSize, true);
                        dynamic size = null;
                        if (findSize == connect.Справочники.ор_Размеры.ПустаяСсылка())
                        {
                            // необходимо добавить тип
                            size = connect.Справочники.ор_Размеры.СоздатьЭлемент();
                            size.Наименование = listBrandsShop[i].listSizes[j].strSize;
                            size.Записать();
                            findSize = size.Ссылка;
                        }
                        // ------------------------

                        dynamic new_row = brand.СоответствиеРазмеров.Добавить();
                        new_row.Размер = findSize;
                        new_row.РусскийРазмер = listBrandsShop[i].listSizes[j].iRussianSize;

                        if (new_row != null) Marshal.FinalReleaseComObject(new_row);
                        if (size != null) Marshal.FinalReleaseComObject(size);
                        if (findSize != null) Marshal.FinalReleaseComObject(findSize);
                    }

                    brand.Записать();

                    if (brand != null) Marshal.FinalReleaseComObject(brand);
                    if (find != null) Marshal.FinalReleaseComObject(find);

                    backgroundWorker1.ReportProgress(0, (i + 1).ToString() + ". Бренд " + listBrandsShop[i].strName + " успешно синхронизирован.");
                }
                
                if (listBrands.Count > 0 || listBrandsShop.Count > 0)
                    backgroundWorker1.ReportProgress(2, "Бренды успешно синхронизированы.");

                /*for (int i = 0; i < listBrands.Count(); i++)
                {
                    Marshal.FinalReleaseComObject(listBrands[i].ref_);
                }

                for (int i = 0; i < listBrandsShop.Count(); i++)
                {
                    Marshal.FinalReleaseComObject(listBrandsShop[i].ref_);
                }*/

                listBrands.Clear();
                listBrandsShop.Clear();
            }
            catch (Exception ex)
            {
                backgroundWorker1.ReportProgress(1, "Ошибка при синхронизации: " + ex.Message);
                return false;
            }

            return true;
        }

        private bool SyncTovars()
        {
            try
            {
                DateTime date = dictShops[m_strSelectedShop].dateSync;

                if (listTovars.Count() > 0 || listTovarsShop.Count() > 0)
                {
                    backgroundWorker1.ReportProgress(0, "");
                    backgroundWorker1.ReportProgress(0, "Синхронизация номенклатур");

                    bool bExistsCross = false;
                    for (int i = 0; i < listTovars.Count; i++)
                    {
                        for (int j = listTovarsShop.Count - 1; j >= 0; j--)
                        {
                            if (listTovars[i].strCodeSync == listTovarsShop[j].strCodeSync)
                            {
                                //listSuppliersShop.RemoveAt(j);
                                bExistsCross = true;
                            }
                        }
                    }

                    if (bExistsCross)
                    {
                        DialogResult res = MessageBox.Show("Некоторые номенклатуры изменялись и в основной базе и в базе магазина. \nЕсли необходимо брать данные из основной базы - нажмите 'Yes' или 'Да', если из базы магазина - 'No' или 'Нет'", "Предупреждение", MessageBoxButtons.YesNo);
                        if (res == System.Windows.Forms.DialogResult.Yes)
                        {
                            for (int i = 0; i < listTovars.Count; i++)
                            {
                                for (int j = listTovarsShop.Count - 1; j >= 0; j--)
                                {
                                    if (listTovars[i].strCodeSync == listTovarsShop[j].strCodeSync)
                                    {
                                        listTovarsShop.RemoveAt(j);
                                    }
                                }
                            }
                        }
                        else
                        {
                            for (int i = 0; i < listTovarsShop.Count; i++)
                            {
                                for (int j = listTovars.Count - 1; j >= 0; j--)
                                {
                                    if (listTovars[j].strCodeSync == listTovarsShop[i].strCodeSync)
                                    {
                                        listTovars.RemoveAt(j);
                                    }
                                }
                            }
                        }
                    }
                }

                // скидываем данные из основной базы
                for (int i = 0; i < listTovars.Count(); i++)
                {
                    // поиск типа
                    // ------------------------
                    dynamic findType = connectShop.Справочники.ВидыНоменклатуры.НайтиПоРеквизиту("КодДляСинхронизации", listTovars[i].type.strCodeSync);
                    dynamic type = null;
                    if (findType == connectShop.Справочники.ВидыНоменклатуры.ПустаяСсылка())
                    {
                        // необходимо добавить тип
                        type = connectShop.Справочники.ВидыНоменклатуры.СоздатьЭлемент();
                        type.Наименование = listTovars[i].type.strName;
                        type.КодДляСинхронизации = listTovars[i].type.strCodeSync;
                        type.Невидимость = listTovars[i].type.bInvisible;
                        type.ИспользованиеХарактеристик = connectShop.Перечисления.ВариантыВеденияДополнительныхДанныхПоНоменклатуре.ИндивидуальныеДляНоменклатуры;
                        type.ТипНоменклатуры = connectShop.Перечисления.ТипыНоменклатуры.Товар;
                        type.Записать();
                        findType = type.Ссылка;
                    }
                    else
                    {
                        // необходимо изменить тип
                        if (findType.Наименование != listTovars[i].type.strName
                            || findType.Невидимость != listTovars[i].type.bInvisible)
                        {
                            type = findType.ПолучитьОбъект();
                            type.Наименование = listTovars[i].type.strName;
                            type.Невидимость = listTovars[i].type.bInvisible;
                            type.Записать();
                        }
                    }
                    // ------------------------

                    // поиск группы
                    // ------------------------
                    dynamic findGroup = connectShop.Справочники.Номенклатура.НайтиПоРеквизиту("КодДляСинхронизации", listTovars[i].group.strCodeSync);
                    dynamic group = null;
                    if (findGroup == connectShop.Справочники.Номенклатура.ПустаяСсылка())
                    {
                        // поиск бренда
                        // ------------------------
                        dynamic findBrand = connectShop.Справочники.ор_ТорговыеМарки.НайтиПоРеквизиту("КодДляСинхронизации", listTovars[i].group.brand.strCodeSync);
                        dynamic brand = null;
                        if (findBrand == connectShop.Справочники.ор_ТорговыеМарки.ПустаяСсылка())
                        {
                            // необходимо добавить
                            brand = connectShop.Справочники.ор_ТорговыеМарки.СоздатьЭлемент();
                            brand.Наименование = listTovars[i].group.brand.strName;
                            brand.КодДляСинхронизации = listTovars[i].group.brand.strCodeSync;
                            brand.Сортировка = listTovars[i].group.brand.iSortPos;
                            brand.Записать();
                            findBrand = brand.Ссылка;
                        }
                        else
                        {
                            // необходимо изменить
                            if (findBrand.Наименование != listTovars[i].group.brand.strName ||
                                findBrand.Сортировка != listTovars[i].group.brand.iSortPos)
                            {
                                brand = findBrand.ПолучитьОбъект();
                                brand.Наименование = listTovars[i].group.brand.strName;
                                brand.Сортировка = listTovars[i].group.brand.iSortPos;
                                brand.Записать();
                            }
                        }
                        // ------------------------
                        // поиск сезона
                        // ------------------------
                        dynamic findSeason = connectShop.Справочники.ор_Сезоны.НайтиПоРеквизиту("КодДляСинхронизации", listTovars[i].group.collection.strCodeSync);
                        dynamic season = null;
                        if (findSeason == connectShop.Справочники.ор_Сезоны.ПустаяСсылка())
                        {
                            // необходимо добавить
                            season = connectShop.Справочники.ор_Сезоны.СоздатьЭлемент();
                            season.Наименование = listTovars[i].group.collection.strName;
                            season.КодДляСинхронизации = listTovars[i].group.collection.strCodeSync;
                            season.Невидимость = listTovars[i].group.collection.bInvisible;
                            season.Сортировка = listTovars[i].group.collection.iSortPos;
                            season.Записать();
                            findSeason = season.Ссылка;
                        }
                        else
                        {
                            // необходимо изменить
                            if (findSeason.Наименование != listTovars[i].group.collection.strName ||
                                findSeason.Невидимость != listTovars[i].group.collection.bInvisible ||
                                findSeason.Сортировка != listTovars[i].group.collection.iSortPos)
                            {
                                season = findSeason.ПолучитьОбъект();
                                season.Наименование = listTovars[i].group.collection.strName;
                                season.Невидимость = listTovars[i].group.collection.bInvisible;
                                season.Сортировка = listTovars[i].group.collection.iSortPos;
                                season.Записать();
                            }
                        }

                        dynamic findSeasonForSite = connectShop.Справочники.ор_Сезоны.НайтиПоРеквизиту("КодДляСинхронизации", listTovars[i].group.collectionForSite.strCodeSync);
                        dynamic seasonForSite = null;
                        if (findSeasonForSite == connectShop.Справочники.ор_Сезоны.ПустаяСсылка())
                        {
                            // необходимо добавить
                            seasonForSite = connectShop.Справочники.ор_Сезоны.СоздатьЭлемент();
                            seasonForSite.Наименование = listTovars[i].group.collectionForSite.strName;
                            seasonForSite.КодДляСинхронизации = listTovars[i].group.collectionForSite.strCodeSync;
                            seasonForSite.Невидимость = listTovars[i].group.collectionForSite.bInvisible;
                            seasonForSite.Сортировка = listTovars[i].group.collectionForSite.iSortPos;
                            seasonForSite.Записать();
                            findSeasonForSite = seasonForSite.Ссылка;
                        }
                        else
                        {
                            // необходимо изменить
                            if (findSeasonForSite.Наименование != listTovars[i].group.collectionForSite.strName ||
                                findSeasonForSite.Невидимость != listTovars[i].group.collectionForSite.bInvisible ||
                                findSeasonForSite.Сортировка != listTovars[i].group.collectionForSite.iSortPos)
                            {
                                seasonForSite = findSeason.ПолучитьОбъект();
                                seasonForSite.Наименование = listTovars[i].group.collectionForSite.strName;
                                seasonForSite.Невидимость = listTovars[i].group.collectionForSite.bInvisible;
                                seasonForSite.Сортировка = listTovars[i].group.collectionForSite.iSortPos;
                                seasonForSite.Записать();
                            }
                        }
                        // ------------------------

                        // необходимо добавить
                        group = connectShop.Справочники.Номенклатура.СоздатьГруппу();
                        group.ор_ТорговаяМарка = findBrand;
                        group.ор_Сезон = findSeason;
                        group.op_СезонДляСайта = findSeasonForSite;
                        group.Наименование = listTovars[i].group.strName;
                        group.Невидимость = listTovars[i].group.bInvisible;
                        group.ЦенаТолькоДляЗарегистрированныхПользователей = listTovars[i].group.bOnlyForRegistered;
                        group.НоваяКоллекция = listTovars[i].group.bNewCollection;
                        group.СкидкаРозн = listTovars[i].group.fDiscount;
                        group.СкидкаОпт = listTovars[i].group.fDiscountWholesale;
                        group.КодДляСинхронизации = listTovars[i].group.strCodeSync;
                        group.Записать();
                        findGroup = group.Ссылка;

                        if (findBrand != null) Marshal.FinalReleaseComObject(findBrand);
                        if (brand != null) Marshal.FinalReleaseComObject(brand);

                        if (findSeason != null) Marshal.FinalReleaseComObject(findSeason);
                        if (season != null) Marshal.FinalReleaseComObject(season);

                        if (findSeasonForSite != null) Marshal.FinalReleaseComObject(findSeasonForSite);
                        if (seasonForSite != null) Marshal.FinalReleaseComObject(seasonForSite);
                    }
                    else
                    {
                        // необходимо изменить
                        if (findGroup.Наименование != listTovars[i].group.strName ||
                            findGroup.op_СезонДляСайта.Наименование != listTovars[i].group.collectionForSite.strName ||
                            findGroup.Невидимость != listTovars[i].group.bInvisible ||
                            findGroup.ЦенаТолькоДляЗарегистрированныхПользователей != listTovars[i].group.bOnlyForRegistered ||
                            findGroup.НоваяКоллекция != listTovars[i].group.bNewCollection ||
                            findGroup.СкидкаРозн != listTovars[i].group.fDiscount ||
                            findGroup.СкидкаОпт != listTovars[i].group.fDiscountWholesale)
                        {
                            // поиск бренда
                            // ------------------------
                            dynamic findBrand = connectShop.Справочники.ор_ТорговыеМарки.НайтиПоРеквизиту("КодДляСинхронизации", listTovars[i].group.brand.strCodeSync);
                            dynamic brand = null;
                            if (findBrand == connectShop.Справочники.ор_ТорговыеМарки.ПустаяСсылка())
                            {
                                // необходимо добавить
                                brand = connectShop.Справочники.ор_ТорговыеМарки.СоздатьЭлемент();
                                brand.Наименование = listTovars[i].group.brand.strName;
                                brand.КодДляСинхронизации = listTovars[i].group.brand.strCodeSync;
                                brand.Сортировка = listTovars[i].group.brand.iSortPos;
                                brand.Записать();
                                findBrand = brand.Ссылка;
                            }
                            else
                            {
                                // необходимо изменить
                                if (findBrand.Наименование != listTovars[i].group.brand.strName ||
                                    findBrand.Сортировка != listTovars[i].group.brand.iSortPos)
                                {
                                    brand = findBrand.ПолучитьОбъект();
                                    brand.Наименование = listTovars[i].group.brand.strName;
                                    brand.Сортировка = listTovars[i].group.brand.iSortPos;
                                    brand.Записать();
                                }
                            }
                            // ------------------------
                            // поиск сезона
                            // ------------------------
                            dynamic findSeason = connectShop.Справочники.ор_Сезоны.НайтиПоРеквизиту("КодДляСинхронизации", listTovars[i].group.collection.strCodeSync);
                            dynamic season = null;
                            if (findSeason == connectShop.Справочники.ор_Сезоны.ПустаяСсылка())
                            {
                                // необходимо добавить
                                season = connectShop.Справочники.ор_Сезоны.СоздатьЭлемент();
                                season.Наименование = listTovars[i].group.collection.strName;
                                season.КодДляСинхронизации = listTovars[i].group.collection.strCodeSync;
                                season.Невидимость = listTovars[i].group.collection.bInvisible;
                                season.Сортировка = listTovars[i].group.collection.iSortPos;
                                season.Записать();
                                findSeason = season.Ссылка;
                            }
                            else
                            {
                                // необходимо изменить
                                if (findSeason.Наименование != listTovars[i].group.collection.strName ||
                                findSeason.Невидимость != listTovars[i].group.collection.bInvisible ||
                                findSeason.Сортировка != listTovars[i].group.collection.iSortPos)
                                {
                                    season = findSeason.ПолучитьОбъект();
                                    season.Наименование = listTovars[i].group.collection.strName;
                                    season.Невидимость = listTovars[i].group.collection.bInvisible;
                                    season.Сортировка = listTovars[i].group.collection.iSortPos;
                                    season.Записать();
                                }
                            }

                            dynamic findSeasonForSite = connectShop.Справочники.ор_Сезоны.НайтиПоРеквизиту("КодДляСинхронизации", listTovars[i].group.collectionForSite.strCodeSync);
                            dynamic seasonForSite = null;
                            if (findSeasonForSite == connectShop.Справочники.ор_Сезоны.ПустаяСсылка())
                            {
                                // необходимо добавить
                                seasonForSite = connectShop.Справочники.ор_Сезоны.СоздатьЭлемент();
                                seasonForSite.Наименование = listTovars[i].group.collectionForSite.strName;
                                seasonForSite.КодДляСинхронизации = listTovars[i].group.collectionForSite.strCodeSync;
                                seasonForSite.Невидимость = listTovars[i].group.collectionForSite.bInvisible;
                                seasonForSite.Сортировка = listTovars[i].group.collectionForSite.iSortPos;
                                seasonForSite.Записать();
                                findSeasonForSite = seasonForSite.Ссылка;
                            }
                            else
                            {
                                // необходимо изменить
                                if (findSeasonForSite.Наименование != listTovars[i].group.collectionForSite.strName ||
                                    findSeasonForSite.Невидимость != listTovars[i].group.collectionForSite.bInvisible ||
                                    findSeasonForSite.Сортировка != listTovars[i].group.collectionForSite.iSortPos)
                                {
                                    seasonForSite = findSeason.ПолучитьОбъект();
                                    seasonForSite.Наименование = listTovars[i].group.collectionForSite.strName;
                                    seasonForSite.Невидимость = listTovars[i].group.collectionForSite.bInvisible;
                                    seasonForSite.Сортировка = listTovars[i].group.collectionForSite.iSortPos;
                                    seasonForSite.Записать();
                                }
                            }
                            // ------------------------

                            group = findGroup.ПолучитьОбъект();
                            group.Наименование = listTovars[i].group.strName;
                            group.ор_ТорговаяМарка = findBrand;
                            group.ор_Сезон = findSeason;
                            group.op_СезонДляСайта = findSeasonForSite;
                            group.Невидимость = listTovars[i].group.bInvisible;
                            group.ЦенаТолькоДляЗарегистрированныхПользователей = listTovars[i].group.bOnlyForRegistered;
                            group.НоваяКоллекция = listTovars[i].group.bNewCollection;
                            group.СкидкаРозн = listTovars[i].group.fDiscount;
                            group.СкидкаОпт = listTovars[i].group.fDiscountWholesale;
                            group.Записать();

                            if (findBrand != null) Marshal.FinalReleaseComObject(findBrand);
                            if (brand != null) Marshal.FinalReleaseComObject(brand);

                            if (findSeason != null) Marshal.FinalReleaseComObject(findSeason);
                            if (season != null) Marshal.FinalReleaseComObject(season);

                            if (findSeasonForSite != null) Marshal.FinalReleaseComObject(findSeasonForSite);
                            if (seasonForSite != null) Marshal.FinalReleaseComObject(seasonForSite);
                        }
                    }
                    // ------------------------

                    string strCodeSync = listTovars[i].strCodeSync;
                    dynamic find = connectShop.Справочники.Номенклатура.НайтиПоРеквизиту("КодДляСинхронизации", strCodeSync);
                    //dynamic tovar = null;
                    if (find == connectShop.Справочники.Номенклатура.ПустаяСсылка())
                    {
                        // необходимо добавить номенклатуру
                        tovar = connectShop.Справочники.Номенклатура.СоздатьЭлемент();
                    }
                    else
                    {
                        // необходимо изменить номенклатуру
                        tovar = find.ПолучитьОбъект();
                    }

                    tovar.Артикул = listTovars[i].strArticul;
                    tovar.ВидНоменклатуры = findType;
                    tovar.ЕдиницаИзмерения = connectShop.Справочники.БазовыеЕдиницыИзмерения.НайтиПоНаименованию("шт", true);
                    tovar.Наименование = listTovars[i].strName;
                    tovar.НаименованиеПолное = listTovars[i].strFullName;
                    tovar.Родитель = findGroup;
                    tovar.СтавкаНДС = connectShop.Перечисления.СтавкиНДС.БезНДС;
                    tovar.ПометкаУдаления = listTovars[i].bDeleted;

                    tovar.Описание = listTovars[i].strNote;
                    tovar.НаименованиеДляСайта = listTovars[i].strNameForSite;
                    tovar.ПримечаниеДляСайта = listTovars[i].strNoteForSite;
                    tovar.СкидкаРозн = listTovars[i].fDiscount;
                    tovar.ДляПолных = listTovars[i].bBigSize;
                    tovar.Сортировка = listTovars[i].iSort;
                    tovar.Новинка = listTovars[i].bNew;
                    tovar.ЦенаЕвро = listTovars[i].fPriceInEuro;
                    tovar.Невидимость = listTovars[i].bInvisible;

                    tovar.КодДляСинхронизации = listTovars[i].strCodeSync;

                    tovar.Записать();

                    // сохранение изображений
                    if (listTovars[i].dateChangeImages >= date)
                    {
                        dynamic null_pic = connect.Справочники.Файлы.ПустаяСсылка();
                        tovar.ФайлКартинки = null_pic;
                        tovar.ФайлКартинки2 = null_pic;
                        tovar.ФайлКартинки3 = null_pic;
                        tovar.ФайлКартинки4 = null_pic;
                        tovar.ФайлКартинки5 = null_pic;

                        string tmp_dir = Environment.GetEnvironmentVariable("TEMP") + "\\tmp.jpg";
                        // необходимо обновить изображения
                        dynamic data1 = listTovars[i].ref_.ФайлКартинки.ТекущаяВерсия.ФайлХранилище.Получить();
                        if (data1 != null)
                        {
                            data1.Записать(tmp_dir);
                            SavePictureShop(tmp_dir, 1);
                        }
                        dynamic data2 = listTovars[i].ref_.ФайлКартинки2.ТекущаяВерсия.ФайлХранилище.Получить();
                        if (data2 != null)
                        {
                            data2.Записать(tmp_dir);
                            SavePictureShop(tmp_dir, 2);
                        }
                        dynamic data3 = listTovars[i].ref_.ФайлКартинки3.ТекущаяВерсия.ФайлХранилище.Получить();
                        if (data3 != null)
                        {
                            data3.Записать(tmp_dir);
                            SavePictureShop(tmp_dir, 3);
                        }
                        dynamic data4 = listTovars[i].ref_.ФайлКартинки4.ТекущаяВерсия.ФайлХранилище.Получить();
                        if (data4 != null)
                        {
                            data4.Записать(tmp_dir);
                            SavePictureShop(tmp_dir, 4);
                        }
                        dynamic data5 = listTovars[i].ref_.ФайлКартинки5.ТекущаяВерсия.ФайлХранилище.Получить();
                        if (data5 != null)
                        {
                            data5.Записать(tmp_dir);
                            SavePictureShop(tmp_dir, 5);
                        }

                        tovar.Записать();

                        if (null_pic != null) Marshal.FinalReleaseComObject(null_pic);
                    }

                    // необходимо обновить цвета и размеры
                    if (listTovars[i].dateChangeAttribs >= date)
                    {
                        // выбираем данные из основной базы
                        dynamic query = connect.NewObject("Запрос");
                        query.Текст = "ВЫБРАТЬ Код, КодДляСинхронизации, Наименование, ПометкаУдаления, ВтороеНаименование, РусскийРазмер ИЗ Справочник.ХарактеристикиНоменклатуры ГДЕ Владелец = &Владелец";
                        query.УстановитьПараметр("Владелец", listTovars[i].ref_);
                        dynamic attribs = query.Выполнить().Выбрать();

                        while (attribs.Следующий())
                        {
                            AttribInfo ai = new AttribInfo();
                            ai.strCodeSync = attribs.КодДляСинхронизации;
                            ai.strName1 = attribs.Наименование;
                            ai.strName2 = attribs.ВтороеНаименование;
                            ai.bDeleted = attribs.ПометкаУдаления;
                            ai.iRussianSize = attribs.РусскийРазмер;

                            listTovars[i].listAttribs.Add(ai);
                        }

                        // выбираем данные из базы клиента
                        dynamic queryShop = connectShop.NewObject("Запрос");
                        queryShop.Текст = "ВЫБРАТЬ Код, КодДляСинхронизации, Наименование, ПометкаУдаления, ВтороеНаименование, РусскийРазмер ИЗ Справочник.ХарактеристикиНоменклатуры ГДЕ Владелец = &Владелец";
                        queryShop.УстановитьПараметр("Владелец", tovar.Ссылка);
                        dynamic attribsShop = queryShop.Выполнить().Выбрать();
                        while (attribsShop.Следующий())
                        {
                            AttribInfo ai = new AttribInfo();
                            ai.strCodeSync = attribsShop.КодДляСинхронизации;
                            ai.strName1 = attribsShop.Наименование;
                            ai.strName2 = attribsShop.ВтороеНаименование;
                            ai.bDeleted = attribsShop.ПометкаУдаления;
                            ai.iRussianSize = attribsShop.РусскийРазмер;

                            listTovars[i].listAttribsShop.Add(ai);
                        }

                        for (int j = 0; j < listTovars[i].listAttribs.Count; j++)
                        {
                            int iRes = 0;
                            for (int k = 0; k < listTovars[i].listAttribsShop.Count; k++)
                            {
                                if (listTovars[i].listAttribs[j].strCodeSync == listTovars[i].listAttribsShop[k].strCodeSync)
                                {
                                    //iRes = 2;
                                    //break;

                                    if (listTovars[i].listAttribs[j].strName1 == listTovars[i].listAttribsShop[k].strName1 &&
                                        listTovars[i].listAttribs[j].iRussianSize == listTovars[i].listAttribsShop[k].iRussianSize)
                                    {
                                        iRes = 1;

                                        if (listTovars[i].listAttribs[j].bDeleted != listTovars[i].listAttribsShop[k].bDeleted)
                                            iRes = 3;
                                        break;
                                    }
                                    else
                                    {
                                        iRes = 2;
                                        break;
                                    }
                                }
                            }

                            /*for (int k = 0; k < listTovars[i].listAttribsShop.Count; k++)
                            {
                                if (listTovars[i].listAttribs[j].strCodeSync == listTovars[i].listAttribsShop[k].strCodeSync &&
                                    listTovars[i].listAttribs[j].strName1 == listTovars[i].listAttribsShop[k].strName1)
                                {
                                    iRes = 1;

                                    if (listTovars[i].listAttribs[j].bDeleted != listTovars[i].listAttribsShop[k].bDeleted)
                                        iRes = 3;
                                    break;
                                }

                                if (listTovars[i].listAttribs[j].strCodeSync == listTovars[i].listAttribsShop[k].strCodeSync &&
                                    listTovars[i].listAttribs[j].strName1 != listTovars[i].listAttribsShop[k].strName1)
                                {
                                    iRes = 2;
                                    break;
                                }
                            }*/

                            if (iRes == 0) // не нашел
                            {
                                // вытаскиваем из наименования размер и цвет и добавляем если надо их в справочники
                                attrib = connectShop.Справочники.ХарактеристикиНоменклатуры.СоздатьЭлемент();
                                attrib.Владелец = tovar.Ссылка;
                                attrib.Наименование = listTovars[i].listAttribs[j].strName1;
                                attrib.ВтороеНаименование = listTovars[i].listAttribs[j].strName2;
                                attrib.РусскийРазмер = listTovars[i].listAttribs[j].iRussianSize;
                                attrib.КодДляСинхронизации = listTovars[i].listAttribs[j].strCodeSync;
                                attrib.ПометкаУдаления = listTovars[i].listAttribs[j].bDeleted;
                                ExtractAttribsShop(listTovars[i].listAttribs[j].strName1, 0);

                                if (attrib != null) Marshal.FinalReleaseComObject(attrib);
                                attrib = null;

                            }

                            if (iRes == 2) // нашел, но поменялось название
                            {
                                attrib = connectShop.Справочники.ХарактеристикиНоменклатуры.НайтиПоРеквизиту("КодДляСинхронизации", listTovars[i].listAttribs[j].strCodeSync).ПолучитьОбъект();
                                attrib.Наименование = listTovars[i].listAttribs[j].strName1;
                                attrib.ВтороеНаименование = listTovars[i].listAttribs[j].strName2;
                                attrib.РусскийРазмер = listTovars[i].listAttribs[j].iRussianSize;
                                attrib.ПометкаУдаления = listTovars[i].listAttribs[j].bDeleted;
                                ExtractAttribsShop(listTovars[i].listAttribs[j].strName1, 1);

                                if (attrib != null) Marshal.FinalReleaseComObject(attrib);
                                attrib = null;
                            }

                            if (iRes == 3) // нашел, но изменилась пометка на удаление
                            {
                                attrib = connectShop.Справочники.ХарактеристикиНоменклатуры.НайтиПоРеквизиту("КодДляСинхронизации", listTovars[i].listAttribs[j].strCodeSync).ПолучитьОбъект();
                                attrib.ПометкаУдаления = listTovars[i].listAttribs[j].bDeleted;
                                attrib.Записать();

                                if (attrib != null) Marshal.FinalReleaseComObject(attrib);
                                attrib = null;
                            }
                        }

                        if (query != null) Marshal.FinalReleaseComObject(query);
                        if (attribs != null) Marshal.FinalReleaseComObject(attribs);

                        if (queryShop != null) Marshal.FinalReleaseComObject(queryShop);
                        if (attribsShop != null) Marshal.FinalReleaseComObject(attribsShop);
                    }

                    if (tovar != null) Marshal.FinalReleaseComObject(tovar);
                    if (type != null) Marshal.FinalReleaseComObject(type);
                    if (group != null) Marshal.FinalReleaseComObject(group);
                    if (find != null) Marshal.FinalReleaseComObject(find);
                    if (findType != null) Marshal.FinalReleaseComObject(findType);
                    if (findGroup != null) Marshal.FinalReleaseComObject(findGroup);

                    backgroundWorker1.ReportProgress(0, (i + 1).ToString() + ". Номенклатура " + listTovars[i].strName + " успешно синхронизирована.");
                }

                // скидываем данные из базы магазина
                for (int i = 0; i < listTovarsShop.Count(); i++)
                {
                    // поиск типа
                    // ------------------------
                    dynamic findType = connect.Справочники.ВидыНоменклатуры.НайтиПоРеквизиту("КодДляСинхронизации", listTovarsShop[i].type.strCodeSync);
                    dynamic type = null;
                    if (findType == connect.Справочники.ВидыНоменклатуры.ПустаяСсылка())
                    {
                        // необходимо добавить тип
                        type = connect.Справочники.ВидыНоменклатуры.СоздатьЭлемент();
                        type.Наименование = listTovarsShop[i].type.strName;
                        type.КодДляСинхронизации = listTovarsShop[i].type.strCodeSync;
                        type.Невидимость = listTovarsShop[i].type.bInvisible;
                        type.ИспользованиеХарактеристик = connect.Перечисления.ВариантыВеденияДополнительныхДанныхПоНоменклатуре.ИндивидуальныеДляНоменклатуры;
                        type.ТипНоменклатуры = connect.Перечисления.ТипыНоменклатуры.Товар;
                        type.Синхронизация = true;
                        type.Записать();
                        findType = type.Ссылка;
                    }
                    else
                    {
                        // необходимо изменить тип
                        if (findType.Наименование != listTovarsShop[i].type.strName
                            || findType.Невидимость != listTovarsShop[i].type.bInvisible)
                        {
                            type = findType.ПолучитьОбъект();
                            type.Наименование = listTovarsShop[i].type.strName;
                            type.Невидимость = listTovarsShop[i].type.bInvisible;
                            type.Синхронизация = true;
                            type.Записать();
                        }
                    }
                    // ------------------------

                    // поиск группы
                    // ------------------------
                    dynamic findGroup = connect.Справочники.Номенклатура.НайтиПоРеквизиту("КодДляСинхронизации", listTovarsShop[i].group.strCodeSync);
                    dynamic group = null;
                    if (findGroup == connect.Справочники.Номенклатура.ПустаяСсылка())
                    {
                        // поиск бренда
                        // ------------------------
                        dynamic findBrand = connect.Справочники.ор_ТорговыеМарки.НайтиПоРеквизиту("КодДляСинхронизации", listTovarsShop[i].group.brand.strCodeSync);
                        dynamic brand = null;
                        if (findBrand == connect.Справочники.ор_ТорговыеМарки.ПустаяСсылка())
                        {
                            // необходимо добавить
                            brand = connect.Справочники.ор_ТорговыеМарки.СоздатьЭлемент();
                            brand.Наименование = listTovarsShop[i].group.brand.strName;
                            brand.КодДляСинхронизации = listTovarsShop[i].group.brand.strCodeSync;
                            brand.Сортировка = listTovarsShop[i].group.brand.iSortPos;
                            brand.Синхронизация = true;
                            brand.Записать();
                            findBrand = brand.Ссылка;
                        }
                        else
                        {
                            // необходимо изменить
                            if (findBrand.Наименование != listTovarsShop[i].group.brand.strName ||
                                findBrand.Сортировка != listTovarsShop[i].group.brand.iSortPos)
                            {
                                brand = findBrand.ПолучитьОбъект();
                                brand.Наименование = listTovarsShop[i].group.brand.strName;
                                brand.Сортировка = listTovarsShop[i].group.brand.iSortPos;
                                brand.Синхронизация = true;
                                brand.Записать();
                            }
                        }
                        // ------------------------
                        // поиск сезона
                        // ------------------------
                        dynamic findSeason = connect.Справочники.ор_Сезоны.НайтиПоРеквизиту("КодДляСинхронизации", listTovarsShop[i].group.collection.strCodeSync);
                        dynamic season = null;
                        if (findSeason == connect.Справочники.ор_Сезоны.ПустаяСсылка())
                        {
                            // необходимо добавить
                            season = connect.Справочники.ор_Сезоны.СоздатьЭлемент();
                            season.Наименование = listTovarsShop[i].group.collection.strName;
                            season.КодДляСинхронизации = listTovarsShop[i].group.collection.strCodeSync;
                            season.Невидимость = listTovarsShop[i].group.collection.bInvisible;
                            season.Сортировка = listTovarsShop[i].group.collection.iSortPos;
                            season.Синхронизация = true;
                            season.Записать();
                            findSeason = season.Ссылка;
                        }
                        else
                        {
                            // необходимо изменить
                            if (findSeason.Наименование != listTovarsShop[i].group.collection.strName ||
                                findSeason.Невидимость != listTovarsShop[i].group.collection.bInvisible ||
                                findSeason.Сортировка != listTovarsShop[i].group.collection.iSortPos)
                            {
                                season = findSeason.ПолучитьОбъект();
                                season.Наименование = listTovarsShop[i].group.collection.strName;
                                season.Невидимость = listTovarsShop[i].group.collection.bInvisible;
                                season.Сортировка = listTovarsShop[i].group.collection.iSortPos;
                                season.Синхронизация = true;
                                season.Записать();
                            }
                        }

                        dynamic findSeasonForSite = connect.Справочники.ор_Сезоны.НайтиПоРеквизиту("КодДляСинхронизации", listTovarsShop[i].group.collectionForSite.strCodeSync);
                        dynamic seasonForSite = null;
                        if (findSeasonForSite == connect.Справочники.ор_Сезоны.ПустаяСсылка())
                        {
                            // необходимо добавить
                            seasonForSite = connect.Справочники.ор_Сезоны.СоздатьЭлемент();
                            seasonForSite.Наименование = listTovarsShop[i].group.collectionForSite.strName;
                            seasonForSite.КодДляСинхронизации = listTovarsShop[i].group.collectionForSite.strCodeSync;
                            seasonForSite.Невидимость = listTovarsShop[i].group.collectionForSite.bInvisible;
                            seasonForSite.Сортировка = listTovarsShop[i].group.collectionForSite.iSortPos;
                            seasonForSite.Синхронизация = true;
                            seasonForSite.Записать();
                            findSeasonForSite = seasonForSite.Ссылка;
                        }
                        else
                        {
                            // необходимо изменить
                            if (findSeasonForSite.Наименование != listTovarsShop[i].group.collectionForSite.strName ||
                                findSeasonForSite.Невидимость != listTovarsShop[i].group.collectionForSite.bInvisible ||
                                findSeasonForSite.Сортировка != listTovarsShop[i].group.collectionForSite.iSortPos)
                            {
                                seasonForSite = findSeason.ПолучитьОбъект();
                                seasonForSite.Наименование = listTovarsShop[i].group.collectionForSite.strName;
                                seasonForSite.Невидимость = listTovarsShop[i].group.collectionForSite.bInvisible;
                                seasonForSite.Сортировка = listTovarsShop[i].group.collectionForSite.iSortPos;
                                seasonForSite.Синхронизация = true;
                                seasonForSite.Записать();
                            }
                        }
                        // ------------------------

                        // необходимо добавить
                        group = connect.Справочники.Номенклатура.СоздатьГруппу();
                        group.ор_ТорговаяМарка = findBrand;
                        group.ор_Сезон = findSeason;
                        group.op_СезонДляСайта = findSeasonForSite;
                        group.Наименование = listTovarsShop[i].group.strName;
                        group.Невидимость = listTovarsShop[i].group.bInvisible;
                        group.ЦенаТолькоДляЗарегистрированныхПользователей = listTovarsShop[i].group.bOnlyForRegistered;
                        group.НоваяКоллекция = listTovarsShop[i].group.bNewCollection;
                        group.СкидкаРозн = listTovarsShop[i].group.fDiscount;
                        group.СкидкаОпт = listTovarsShop[i].group.fDiscountWholesale;
                        group.КодДляСинхронизации = listTovarsShop[i].group.strCodeSync;
                        group.Синхронизация = true;
                        group.Записать();
                        findGroup = group.Ссылка;

                        if (findBrand != null) Marshal.FinalReleaseComObject(findBrand);
                        if (brand != null) Marshal.FinalReleaseComObject(brand);

                        if (findSeason != null) Marshal.FinalReleaseComObject(findSeason);
                        if (season != null) Marshal.FinalReleaseComObject(season);

                        if (findSeasonForSite != null) Marshal.FinalReleaseComObject(findSeasonForSite);
                        if (seasonForSite != null) Marshal.FinalReleaseComObject(seasonForSite);
                    }
                    else
                    {
                        // необходимо изменить
                        if (findGroup.Наименование != listTovarsShop[i].group.strName ||
                            findGroup.op_СезонДляСайта.Наименование != listTovarsShop[i].group.collectionForSite.strName ||
                            findGroup.Невидимость != listTovarsShop[i].group.bInvisible ||
                            findGroup.ЦенаТолькоДляЗарегистрированныхПользователей != listTovarsShop[i].group.bOnlyForRegistered ||
                            findGroup.НоваяКоллекция != listTovarsShop[i].group.bNewCollection ||
                            findGroup.СкидкаРозн != listTovarsShop[i].group.fDiscount ||
                            findGroup.СкидкаОпт != listTovarsShop[i].group.fDiscountWholesale)
                        {
                            // поиск бренда
                            // ------------------------
                            dynamic findBrand = connect.Справочники.ор_ТорговыеМарки.НайтиПоРеквизиту("КодДляСинхронизации", listTovarsShop[i].group.brand.strCodeSync);
                            dynamic brand = null;
                            if (findBrand == connect.Справочники.ор_ТорговыеМарки.ПустаяСсылка())
                            {
                                // необходимо добавить
                                brand = connect.Справочники.ор_ТорговыеМарки.СоздатьЭлемент();
                                brand.Наименование = listTovarsShop[i].group.brand.strName;
                                brand.КодДляСинхронизации = listTovarsShop[i].group.brand.strCodeSync;
                                brand.Сортировка = listTovarsShop[i].group.brand.iSortPos;
                                brand.Синхронизация = true;
                                brand.Записать();
                                findBrand = brand.Ссылка;
                            }
                            else
                            {
                                // необходимо изменить
                                if (findBrand.Наименование != listTovarsShop[i].group.brand.strName ||
                                    findBrand.Сортировка != listTovarsShop[i].group.brand.iSortPos)
                                {
                                    brand = findBrand.ПолучитьОбъект();
                                    brand.Наименование = listTovarsShop[i].group.brand.strName;
                                    brand.Сортировка = listTovarsShop[i].group.brand.iSortPos;
                                    brand.Синхронизация = true;
                                    brand.Записать();
                                }
                            }
                            // ------------------------
                            // поиск сезона
                            // ------------------------
                            dynamic findSeason = connect.Справочники.ор_Сезоны.НайтиПоРеквизиту("КодДляСинхронизации", listTovarsShop[i].group.collection.strCodeSync);
                            dynamic season = null;
                            if (findSeason == connect.Справочники.ор_Сезоны.ПустаяСсылка())
                            {
                                // необходимо добавить
                                season = connect.Справочники.ор_Сезоны.СоздатьЭлемент();
                                season.Наименование = listTovarsShop[i].group.collection.strName;
                                season.КодДляСинхронизации = listTovarsShop[i].group.collection.strCodeSync;
                                season.Невидимость = listTovarsShop[i].group.collection.bInvisible;
                                season.Сортировка = listTovarsShop[i].group.collection.iSortPos;
                                season.Синхронизация = true;
                                season.Записать();
                                findSeason = season.Ссылка;
                            }
                            else
                            {
                                // необходимо изменить
                                if (findSeason.Наименование != listTovarsShop[i].group.collection.strName ||
                                findSeason.Невидимость != listTovarsShop[i].group.collection.bInvisible ||
                                findSeason.Сортировка != listTovarsShop[i].group.collection.iSortPos)
                                {
                                    season = findSeason.ПолучитьОбъект();
                                    season.Наименование = listTovarsShop[i].group.collection.strName;
                                    season.Невидимость = listTovarsShop[i].group.collection.bInvisible;
                                    season.Сортировка = listTovarsShop[i].group.collection.iSortPos;
                                    season.Синхронизация = true;
                                    season.Записать();
                                }
                            }

                            dynamic findSeasonForSite = connect.Справочники.ор_Сезоны.НайтиПоРеквизиту("КодДляСинхронизации", listTovarsShop[i].group.collectionForSite.strCodeSync);
                            dynamic seasonForSite = null;
                            if (findSeasonForSite == connect.Справочники.ор_Сезоны.ПустаяСсылка())
                            {
                                // необходимо добавить
                                seasonForSite = connect.Справочники.ор_Сезоны.СоздатьЭлемент();
                                seasonForSite.Наименование = listTovarsShop[i].group.collectionForSite.strName;
                                seasonForSite.КодДляСинхронизации = listTovarsShop[i].group.collectionForSite.strCodeSync;
                                seasonForSite.Невидимость = listTovarsShop[i].group.collectionForSite.bInvisible;
                                seasonForSite.Сортировка = listTovarsShop[i].group.collectionForSite.iSortPos;
                                seasonForSite.Синхронизация = true;
                                seasonForSite.Записать();
                                findSeasonForSite = seasonForSite.Ссылка;
                            }
                            else
                            {
                                // необходимо изменить
                                if (findSeasonForSite.Наименование != listTovarsShop[i].group.collectionForSite.strName ||
                                    findSeasonForSite.Невидимость != listTovarsShop[i].group.collectionForSite.bInvisible ||
                                    findSeasonForSite.Сортировка != listTovarsShop[i].group.collectionForSite.iSortPos)
                                {
                                    seasonForSite = findSeason.ПолучитьОбъект();
                                    seasonForSite.Наименование = listTovarsShop[i].group.collectionForSite.strName;
                                    seasonForSite.Невидимость = listTovarsShop[i].group.collectionForSite.bInvisible;
                                    seasonForSite.Сортировка = listTovarsShop[i].group.collectionForSite.iSortPos;
                                    //seasonForSite.Синхронизация = true;
                                    seasonForSite.Записать();
                                }
                            }
                            // ------------------------

                            group = findGroup.ПолучитьОбъект();
                            group.Наименование = listTovarsShop[i].group.strName;
                            group.ор_ТорговаяМарка = findBrand;
                            group.ор_Сезон = findSeason;
                            group.op_СезонДляСайта = findSeasonForSite;
                            group.Невидимость = listTovarsShop[i].group.bInvisible;
                            group.ЦенаТолькоДляЗарегистрированныхПользователей = listTovarsShop[i].group.bOnlyForRegistered;
                            group.НоваяКоллекция = listTovarsShop[i].group.bNewCollection;
                            group.СкидкаРозн = listTovarsShop[i].group.fDiscount;
                            group.СкидкаОпт = listTovarsShop[i].group.fDiscountWholesale;
                            group.Синхронизация = true;
                            group.Записать();

                            if (findBrand != null) Marshal.FinalReleaseComObject(findBrand);
                            if (brand != null) Marshal.FinalReleaseComObject(brand);

                            if (findSeason != null) Marshal.FinalReleaseComObject(findSeason);
                            if (season != null) Marshal.FinalReleaseComObject(season);

                            if (findSeasonForSite != null) Marshal.FinalReleaseComObject(findSeasonForSite);
                            if (seasonForSite != null) Marshal.FinalReleaseComObject(seasonForSite);
                        }
                    }
                    // ------------------------

                    string strCodeSync = listTovarsShop[i].strCodeSync;
                    dynamic find = connect.Справочники.Номенклатура.НайтиПоРеквизиту("КодДляСинхронизации", strCodeSync);
                    //dynamic tovar = null;
                    if (find == connect.Справочники.Номенклатура.ПустаяСсылка())
                    {
                        // необходимо добавить номенклатуру
                        tovar = connect.Справочники.Номенклатура.СоздатьЭлемент();
                    }
                    else
                    {
                        // необходимо изменить номенклатуру
                        tovar = find.ПолучитьОбъект();
                    }

                    tovar.Артикул = listTovarsShop[i].strArticul;
                    tovar.ВидНоменклатуры = findType;
                    tovar.ЕдиницаИзмерения = connect.Справочники.БазовыеЕдиницыИзмерения.НайтиПоНаименованию("шт", true);
                    tovar.Наименование = listTovarsShop[i].strName;
                    tovar.НаименованиеПолное = listTovarsShop[i].strFullName;
                    tovar.Родитель = findGroup;
                    tovar.СтавкаНДС = connect.Перечисления.СтавкиНДС.БезНДС;
                    tovar.ПометкаУдаления = listTovarsShop[i].bDeleted;

                    tovar.Описание = listTovarsShop[i].strNote;
                    tovar.НаименованиеДляСайта = listTovarsShop[i].strNameForSite;
                    tovar.ПримечаниеДляСайта = listTovarsShop[i].strNoteForSite;
                    tovar.СкидкаРозн = listTovarsShop[i].fDiscount;
                    tovar.ДляПолных = listTovarsShop[i].bBigSize;
                    tovar.Сортировка = listTovarsShop[i].iSort;
                    tovar.Новинка = listTovarsShop[i].bNew;
                    tovar.ЦенаЕвро = listTovarsShop[i].fPriceInEuro;
                    tovar.Невидимость = listTovarsShop[i].bInvisible;

                    tovar.КодДляСинхронизации = listTovarsShop[i].strCodeSync;

                    tovar.Синхронизация = true;
                    if (listTovarsShop[i].dateChangeImages >= date)
                        tovar.СинхронизацияИзображений = true;
                    if (listTovarsShop[i].dateChangeAttribs >= date)
                        tovar.СинхронизацияАттрибутов = true;

                    tovar.Записать();

                    // сохранение изображений
                    if (listTovarsShop[i].dateChangeImages >= date)
                    {
                        dynamic null_pic = connectShop.Справочники.Файлы.ПустаяСсылка();
                        tovar.ФайлКартинки = null_pic;
                        tovar.ФайлКартинки2 = null_pic;
                        tovar.ФайлКартинки3 = null_pic;
                        tovar.ФайлКартинки4 = null_pic;
                        tovar.ФайлКартинки5 = null_pic;

                        string tmp_dir = Environment.GetEnvironmentVariable("TEMP") + "\\tmp.jpg";
                        // необходимо обновить изображения
                        dynamic data1 = listTovarsShop[i].ref_.ФайлКартинки.ТекущаяВерсия.ФайлХранилище.Получить();
                        if (data1 != null)
                        {
                            data1.Записать(tmp_dir);
                            SavePicture(tmp_dir, 1);
                        }
                        dynamic data2 = listTovarsShop[i].ref_.ФайлКартинки2.ТекущаяВерсия.ФайлХранилище.Получить();
                        if (data2 != null)
                        {
                            data2.Записать(tmp_dir);
                            SavePicture(tmp_dir, 2);
                        }
                        dynamic data3 = listTovarsShop[i].ref_.ФайлКартинки3.ТекущаяВерсия.ФайлХранилище.Получить();
                        if (data3 != null)
                        {
                            data3.Записать(tmp_dir);
                            SavePicture(tmp_dir, 3);
                        }
                        dynamic data4 = listTovarsShop[i].ref_.ФайлКартинки4.ТекущаяВерсия.ФайлХранилище.Получить();
                        if (data4 != null)
                        {
                            data4.Записать(tmp_dir);
                            SavePicture(tmp_dir, 4);
                        }
                        dynamic data5 = listTovarsShop[i].ref_.ФайлКартинки5.ТекущаяВерсия.ФайлХранилище.Получить();
                        if (data5 != null)
                        {
                            data5.Записать(tmp_dir);
                            SavePicture(tmp_dir, 5);
                        }

                        tovar.Записать();

                        if (null_pic != null) Marshal.FinalReleaseComObject(null_pic);
                    }

                    // необходимо обновить цвета и размеры
                    if (listTovarsShop[i].dateChangeAttribs >= date)
                    {
                        // выбираем данные из основной базы
                        dynamic query = connectShop.NewObject("Запрос");
                        query.Текст = "ВЫБРАТЬ Код, КодДляСинхронизации, Наименование, ВтороеНаименование, ПометкаУдаления, РусскийРазмер ИЗ Справочник.ХарактеристикиНоменклатуры ГДЕ Владелец = &Владелец";
                        query.УстановитьПараметр("Владелец", listTovarsShop[i].ref_);
                        dynamic attribs = query.Выполнить().Выбрать();

                        while (attribs.Следующий())
                        {
                            AttribInfo ai = new AttribInfo();
                            ai.strCodeSync = attribs.КодДляСинхронизации;
                            ai.strName1 = attribs.Наименование;
                            ai.strName2 = attribs.ВтороеНаименование;
                            ai.bDeleted = attribs.ПометкаУдаления;
                            ai.iRussianSize = attribs.РусскийРазмер;

                            listTovarsShop[i].listAttribs.Add(ai);
                        }

                        // выбираем данные из базы клиента
                        dynamic queryShop = connect.NewObject("Запрос");
                        queryShop.Текст = "ВЫБРАТЬ Код, КодДляСинхронизации, Наименование, ВтороеНаименование, ПометкаУдаления, РусскийРазмер ИЗ Справочник.ХарактеристикиНоменклатуры ГДЕ Владелец = &Владелец";
                        queryShop.УстановитьПараметр("Владелец", tovar.Ссылка);
                        dynamic attribsShop = queryShop.Выполнить().Выбрать();
                        while (attribsShop.Следующий())
                        {
                            AttribInfo ai = new AttribInfo();
                            ai.strCodeSync = attribsShop.КодДляСинхронизации;
                            ai.strName1 = attribsShop.Наименование;
                            ai.strName2 = attribsShop.ВтороеНаименование;
                            ai.bDeleted = attribsShop.ПометкаУдаления;
                            ai.iRussianSize = attribsShop.РусскийРазмер;

                            listTovarsShop[i].listAttribsShop.Add(ai);
                        }

                        for (int j = 0; j < listTovarsShop[i].listAttribs.Count; j++)
                        {
                            int iRes = 0;
                            for (int k = 0; k < listTovarsShop[i].listAttribsShop.Count; k++)
                            {
                                if (listTovarsShop[i].listAttribs[j].strCodeSync == listTovarsShop[i].listAttribsShop[k].strCodeSync)
                                {
                                    //iRes = 2;
                                    //break;

                                    if (listTovarsShop[i].listAttribs[j].strName1 == listTovarsShop[i].listAttribsShop[k].strName1 &&
                                        listTovarsShop[i].listAttribs[j].iRussianSize == listTovarsShop[i].listAttribsShop[k].iRussianSize)
                                    {
                                        iRes = 1;

                                        if (listTovarsShop[i].listAttribs[j].bDeleted != listTovarsShop[i].listAttribsShop[k].bDeleted)
                                            iRes = 3;
                                        break;
                                    }
                                    else
                                    {
                                        iRes = 2;
                                        break;
                                    }
                                }
                            }

                            if (iRes == 0) // не нашел
                            {
                                // вытаскиваем из наименования размер и цвет и добавляем если надо их в справочники
                                attrib = connect.Справочники.ХарактеристикиНоменклатуры.СоздатьЭлемент();
                                attrib.Владелец = tovar.Ссылка;
                                attrib.Наименование = listTovarsShop[i].listAttribs[j].strName1;
                                attrib.ВтороеНаименование = listTovarsShop[i].listAttribs[j].strName2;
                                attrib.РусскийРазмер = listTovarsShop[i].listAttribs[j].iRussianSize;
                                attrib.КодДляСинхронизации = listTovarsShop[i].listAttribs[j].strCodeSync;
                                attrib.ПометкаУдаления = listTovarsShop[i].listAttribs[j].bDeleted;
                                ExtractAttribs(listTovarsShop[i].listAttribs[j].strName1, 0);

                                if (attrib != null) Marshal.FinalReleaseComObject(attrib);
                                attrib = null;

                            }

                            if (iRes == 2) // нашел, но поменялось название
                            {
                                attrib = connect.Справочники.ХарактеристикиНоменклатуры.НайтиПоРеквизиту("КодДляСинхронизации", listTovarsShop[i].listAttribs[j].strCodeSync).ПолучитьОбъект();
                                attrib.Наименование = listTovarsShop[i].listAttribs[j].strName1;
                                attrib.ВтороеНаименование = listTovarsShop[i].listAttribs[j].strName2;
                                attrib.РусскийРазмер = listTovarsShop[i].listAttribs[j].iRussianSize;
                                attrib.ПометкаУдаления = listTovarsShop[i].listAttribs[j].bDeleted;
                                ExtractAttribs(listTovarsShop[i].listAttribs[j].strName1, 1);

                                if (attrib != null) Marshal.FinalReleaseComObject(attrib);
                                attrib = null;
                            }

                            if (iRes == 3) // нашел, но изменилась пометка на удаление
                            {
                                attrib = connect.Справочники.ХарактеристикиНоменклатуры.НайтиПоРеквизиту("КодДляСинхронизации", listTovarsShop[i].listAttribs[j].strCodeSync).ПолучитьОбъект();
                                attrib.ПометкаУдаления = listTovarsShop[i].listAttribs[j].bDeleted;
                                attrib.Записать();

                                if (attrib != null) Marshal.FinalReleaseComObject(attrib);
                                attrib = null;
                            }
                        }

                        if (query != null) Marshal.FinalReleaseComObject(query);
                        if (attribs != null) Marshal.FinalReleaseComObject(attribs);

                        if (queryShop != null) Marshal.FinalReleaseComObject(queryShop);
                        if (attribsShop != null) Marshal.FinalReleaseComObject(attribsShop);
                    }

                    if (tovar != null) Marshal.FinalReleaseComObject(tovar);
                    if (type != null) Marshal.FinalReleaseComObject(type);
                    if (group != null) Marshal.FinalReleaseComObject(group);
                    if (find != null) Marshal.FinalReleaseComObject(find);
                    if (findType != null) Marshal.FinalReleaseComObject(findType);
                    if (findGroup != null) Marshal.FinalReleaseComObject(findGroup);

                    backgroundWorker1.ReportProgress(0, (i + 1).ToString() + ". Номенклатура " + listTovarsShop[i].strName + " успешно синхронизирована.");
                }

                if (listTovars.Count > 0 || listTovarsShop.Count > 0)
                    backgroundWorker1.ReportProgress(2, "Номенклатуры успешно синхронизированы.");

                for (int i = 0; i < listTovars.Count(); i++)
                {
                    Marshal.FinalReleaseComObject(listTovars[i].ref_);
                }

                for (int i = 0; i < listTovarsShop.Count(); i++)
                {
                    Marshal.FinalReleaseComObject(listTovarsShop[i].ref_);
                }

                listTovars.Clear();
                listTovarsShop.Clear();
            }
            catch (Exception ex)
            {
                backgroundWorker1.ReportProgress(1, "Ошибка при синхронизации: " + ex.Message);
                return false;
            }

            return true;
        }

        private bool SyncPrice()
        {
            try
            {
                //DateTime date = dictShops[m_strSelectedShop].dateSyncTovars;

                if (listPriceDocuments.Count() > 0 || listPriceDocumentsShop.Count() > 0)
                {
                    backgroundWorker1.ReportProgress(0, "");
                    backgroundWorker1.ReportProgress(0, "Синхронизация документов установки цен");

                    bool bExistsCross = false;
                    for (int i = 0; i < listPriceDocuments.Count; i++)
                    {
                        for (int j = listPriceDocumentsShop.Count - 1; j >= 0; j--)
                        {
                            if (listPriceDocuments[i].strCodeSync == listPriceDocumentsShop[j].strCodeSync)
                            {
                                //listSuppliersShop.RemoveAt(j);
                                bExistsCross = true;
                            }
                        }
                    }

                    if (bExistsCross)
                    {
                        DialogResult res = MessageBox.Show("Некоторые документы установки цены изменялись и в основной базе и в базе магазина. \nЕсли необходимо брать данные из основной базы - нажмите 'Yes' или 'Да', если из базы магазина - 'No' или 'Нет'", "Предупреждение", MessageBoxButtons.YesNo);
                        if (res == System.Windows.Forms.DialogResult.Yes)
                        {
                            for (int i = 0; i < listPriceDocuments.Count; i++)
                            {
                                for (int j = listPriceDocumentsShop.Count - 1; j >= 0; j--)
                                {
                                    if (listPriceDocuments[i].strCodeSync == listPriceDocumentsShop[j].strCodeSync)
                                    {
                                        listPriceDocumentsShop.RemoveAt(j);
                                    }
                                }
                            }
                        }
                        else
                        {
                            for (int i = 0; i < listPriceDocumentsShop.Count; i++)
                            {
                                for (int j = listPriceDocuments.Count - 1; j >= 0; j--)
                                {
                                    if (listPriceDocuments[j].strCodeSync == listPriceDocumentsShop[i].strCodeSync)
                                    {
                                        listPriceDocuments.RemoveAt(j);
                                    }
                                }
                            }
                        }
                    }
                }

                // передаем данные по ценам
                for (int i = 0; i < listPriceDocuments.Count(); i++)
                {

                    // если не находим документа с таким номером и датой, просто добавляем его и проводим
                    // если находим, тогда добавляем только позиции
                    //bool bProvodka = false;
                    //bool bDeleted = false;

                    dynamic findDoc = connectShop.Документы.УстановкаЦенНоменклатуры.НайтиПоРеквизиту("КодДляСинхронизации", listPriceDocuments[i].strCodeSync);
                    dynamic doc = null;
                    if (findDoc == connectShop.Документы.УстановкаЦенНоменклатуры.ПустаяСсылка())
                    {
                        // необходимо добавить номенклатуру
                        doc = connectShop.Документы.УстановкаЦенНоменклатуры.СоздатьДокумент();
                    }
                    else
                    {
                        // необходимо изменить номенклатуру
                        doc = findDoc.ПолучитьОбъект();
                    }

                    dynamic type_price = connectShop.Справочники.ВидыЦен.НайтиПоНаименованию("Вид цены", true);
                    if (type_price == connectShop.Справочники.ВидыЦен.ПустаяСсылка())
                    {
                        dynamic type_price_object = connectShop.Справочники.ВидыЦен.СоздатьЭлемент();
                        type_price_object.Наименование = "Вид цены";
                        type_price_object.ИспользоватьПриПродаже = true;
                        type_price_object.Записать();
                        type_price = type_price_object.Ссылка;

                        if (type_price_object != null) Marshal.FinalReleaseComObject(type_price_object);
                    }

                    dynamic user = connectShop.Справочники.Пользователи.НайтиПоНаименованию("<Не указан>", true);

                    doc.КодДляСинхронизации = listPriceDocuments[i].strCodeSync;
                    doc.Номер = listPriceDocuments[i].strCode;
                    doc.Дата = listPriceDocuments[i].date;
                    doc.Комментарий = listPriceDocuments[i].strNote;
                    if (listPriceDocuments[i].bDeleted) doc.ПометкаУдаления = true;
                    doc.Ответственный = user.Ссылка;
                    doc.Товары.Очистить();
                    doc.ВидыЦен.Очистить();

                    dynamic row_type_price = null;
                    row_type_price = doc.ВидыЦен.Добавить();
                    row_type_price.ВидЦены = type_price.Ссылка;

                    // добавление товаров
                    for (int j = 0; j < listPriceDocuments[i].listTovars.Count(); j++)
                    {
                        dynamic findTovar = connectShop.Справочники.Номенклатура.НайтиПоРеквизиту("КодДляСинхронизации", listPriceDocuments[i].listTovars[j].strTovarSyncCode);
                        if (findTovar == connectShop.Справочники.Номенклатура.ПустаяСсылка())
                        {
                            if (findTovar != null) Marshal.FinalReleaseComObject(findTovar);
                            continue;
                        }
                        dynamic findAttrib = null;
                        if (listPriceDocuments[i].listTovars[j].strAttribSyncCode != "")
                        {
                            findAttrib = connectShop.Справочники.ХарактеристикиНоменклатуры.НайтиПоРеквизиту("КодДляСинхронизации", listPriceDocuments[i].listTovars[j].strAttribSyncCode);
                            if (findAttrib == connectShop.Справочники.ХарактеристикиНоменклатуры.ПустаяСсылка())
                            {
                                if (findTovar != null) Marshal.FinalReleaseComObject(findTovar);
                                continue;
                            }
                        }

                        dynamic new_row = doc.Товары.Добавить();
                        new_row.Номенклатура = findTovar;
                        new_row.ВидЦены = type_price.Ссылка;
                        new_row.Цена = listPriceDocuments[i].listTovars[j].price;
                        if (listPriceDocuments[i].listTovars[j].strAttribSyncCode == "")
                            new_row.Характеристика = connectShop.Справочники.ХарактеристикиНоменклатуры.ПустаяСсылка();
                        else
                            new_row.Характеристика = findAttrib;

                        if (new_row != null) Marshal.FinalReleaseComObject(new_row);
                        if (findTovar != null) Marshal.FinalReleaseComObject(findTovar);
                        if (findAttrib != null) Marshal.FinalReleaseComObject(findAttrib);
                    }

                    if (listPriceDocuments[i].bProvodka)
                        doc.Записать(connectShop.РежимЗаписиДокумента.Проведение);
                    else
                        doc.Записать(connectShop.РежимЗаписиДокумента.ОтменаПроведения);

                    if (type_price != null) Marshal.FinalReleaseComObject(type_price);
                    if (user != null) Marshal.FinalReleaseComObject(user);
                    if (findDoc != null) Marshal.FinalReleaseComObject(findDoc);
                    //if (queryPD != null) Marshal.FinalReleaseComObject(queryPD);
                    //if (objectsPD != null) Marshal.FinalReleaseComObject(objectsPD);
                    if (doc != null) Marshal.FinalReleaseComObject(doc);
                    if (row_type_price != null) Marshal.FinalReleaseComObject(row_type_price);

                    backgroundWorker1.ReportProgress(0, (i + 1).ToString() + ". Документ установки цены №" + listPriceDocuments[i].strCode + " от " + listPriceDocuments[i].date.ToShortDateString() + " успешно синхронизирован.");
                }

                // передаем данные по ценам
                for (int i = 0; i < listPriceDocumentsShop.Count(); i++)
                {

                    // если не находим документа с таким номером и датой, просто добавляем его и проводим
                    // если находим, тогда добавляем только позиции
                    //bool bProvodka = false;
                    //bool bDeleted = false;

                    dynamic findDoc = connect.Документы.УстановкаЦенНоменклатуры.НайтиПоРеквизиту("КодДляСинхронизации", listPriceDocumentsShop[i].strCodeSync);
                    dynamic doc = null;
                    if (findDoc == connect.Документы.УстановкаЦенНоменклатуры.ПустаяСсылка())
                    {
                        // необходимо добавить номенклатуру
                        doc = connect.Документы.УстановкаЦенНоменклатуры.СоздатьДокумент();
                    }
                    else
                    {
                        // необходимо изменить номенклатуру
                        doc = findDoc.ПолучитьОбъект();
                    }

                    dynamic type_price = connect.Справочники.ВидыЦен.НайтиПоНаименованию("Вид цены", true);
                    if (type_price == connect.Справочники.ВидыЦен.ПустаяСсылка())
                    {
                        dynamic type_price_object = connect.Справочники.ВидыЦен.СоздатьЭлемент();
                        type_price_object.Наименование = "Вид цены";
                        type_price_object.ИспользоватьПриПродаже = true;
                        type_price_object.Записать();
                        type_price = type_price_object.Ссылка;

                        if (type_price_object != null) Marshal.FinalReleaseComObject(type_price_object);
                    }

                    dynamic user = connect.Справочники.Пользователи.НайтиПоНаименованию("<Не указан>", true);

                    doc.КодДляСинхронизации = listPriceDocumentsShop[i].strCodeSync;
                    doc.Номер = listPriceDocumentsShop[i].strCode;
                    doc.Дата = listPriceDocumentsShop[i].date;
                    doc.Комментарий = listPriceDocumentsShop[i].strNote;
                    if (listPriceDocumentsShop[i].bDeleted) doc.ПометкаУдаления = true;
                    doc.Ответственный = user.Ссылка;
                    doc.Товары.Очистить();
                    doc.ВидыЦен.Очистить();

                    dynamic row_type_price = null;
                    row_type_price = doc.ВидыЦен.Добавить();
                    row_type_price.ВидЦены = type_price.Ссылка;

                    // добавление товаров
                    for (int j = 0; j < listPriceDocumentsShop[i].listTovars.Count(); j++)
                    {
                        dynamic findTovar = connect.Справочники.Номенклатура.НайтиПоРеквизиту("КодДляСинхронизации", listPriceDocumentsShop[i].listTovars[j].strTovarSyncCode);
                        if (findTovar == connect.Справочники.Номенклатура.ПустаяСсылка())
                        {
                            continue;
                        }
                        dynamic findAttrib = null;
                        if (listPriceDocumentsShop[i].listTovars[j].strAttribSyncCode != "")
                        {
                            findAttrib = connect.Справочники.ХарактеристикиНоменклатуры.НайтиПоРеквизиту("КодДляСинхронизации", listPriceDocumentsShop[i].listTovars[j].strAttribSyncCode);
                            if (findAttrib == connect.Справочники.ХарактеристикиНоменклатуры.ПустаяСсылка())
                            {
                                continue;
                            }
                        }

                        dynamic new_row = doc.Товары.Добавить();
                        new_row.Номенклатура = findTovar;
                        new_row.ВидЦены = type_price.Ссылка;
                        new_row.Цена = listPriceDocumentsShop[i].listTovars[j].price;
                        if (listPriceDocumentsShop[i].listTovars[j].strAttribSyncCode == "")
                            new_row.Характеристика = connect.Справочники.ХарактеристикиНоменклатуры.ПустаяСсылка();
                        else
                            new_row.Характеристика = findAttrib;

                        if (new_row != null) Marshal.FinalReleaseComObject(new_row);
                        if (findTovar != null) Marshal.FinalReleaseComObject(findTovar);
                        if (findAttrib != null) Marshal.FinalReleaseComObject(findAttrib);
                    }

                    if (listPriceDocumentsShop[i].bProvodka)
                        doc.Записать(connect.РежимЗаписиДокумента.Проведение);
                    else
                        doc.Записать(connect.РежимЗаписиДокумента.ОтменаПроведения);

                    if (type_price != null) Marshal.FinalReleaseComObject(type_price);
                    if (user != null) Marshal.FinalReleaseComObject(user);
                    if (findDoc != null) Marshal.FinalReleaseComObject(findDoc);
                    //if (queryPD != null) Marshal.FinalReleaseComObject(queryPD);
                    //if (objectsPD != null) Marshal.FinalReleaseComObject(objectsPD);
                    if (doc != null) Marshal.FinalReleaseComObject(doc);
                    if (row_type_price != null) Marshal.FinalReleaseComObject(row_type_price);

                    backgroundWorker1.ReportProgress(0, (i + 1).ToString() + ". Документ установки цены №" + listPriceDocumentsShop[i].strCode + " от " + listPriceDocumentsShop[i].date.ToShortDateString() + " успешно синхронизирован.");
                }

                if (listPriceDocuments.Count > 0 || listPriceDocumentsShop.Count > 0)
                    backgroundWorker1.ReportProgress(2, "Документы установки цен успешно синхронизированы.");

                listPriceDocuments.Clear();
                listPriceDocumentsShop.Clear();

            }
            catch (Exception ex)
            {
                backgroundWorker1.ReportProgress(1, "Ошибка при синхронизации: " + ex.Message);
                return false;
            }

            return true;
        }

        private bool SyncDiscount()
        {
            try
            {
                //DateTime date = dictShops[m_strSelectedShop].dateSyncTovars;

                // сравниваем акции
                if (ActionDocument == null && ActionShopDocument != null ||
                    ActionDocument != null && ActionShopDocument == null ||
                    ActionDocument != null && ActionShopDocument != null &&
                    (ActionDocument.bDeleted != ActionShopDocument.bDeleted ||
                    ActionDocument.bProvodka != ActionShopDocument.bProvodka ||
                    ActionDocument.dateBegin != ActionShopDocument.dateBegin ||
                    ActionDocument.dateEnd != ActionShopDocument.dateEnd ||
                    ActionDocument.strName != ActionShopDocument.strName ||
                    ActionDocument.listDiscounts.Count != ActionShopDocument.listDiscounts.Count
                    || !EqualActionDiscounts()))
                {
                    // нужно синхронизировать
                    backgroundWorker1.ReportProgress(0, "");
                    backgroundWorker1.ReportProgress(0, "Синхронизация маркетинговых акций и скидок.");

                    if (ActionDocument == null && ActionShopDocument != null)
                    {
                        if (ActionShopDocument.ref_ != null)
                            ActionShopDocument.ref_.Удалить();
                        //ActionShopDocument.ref_.ПометкаУдаления = true;
                        //ActionShopDocument.ref_.Записать(connectShop.РежимЗаписиДокумента.ОтменаПроведения);
                    }
                    else
                    {
                        if (ActionDocument != null && ActionShopDocument == null)
                        {
                            ActionShopDocument = new ActionInfo();
                            ActionShopDocument.ref_ = connectShop.Документы.МаркетинговаяАкция.СоздатьДокумент();
                        }

                        ActionShopDocument.ref_.ПометкаУдаления = ActionDocument.bDeleted;
                        ActionShopDocument.ref_.Дата = DateTime.Now;
                        ActionShopDocument.ref_.ДатаНачалаДействия = ActionDocument.dateBegin;
                        ActionShopDocument.ref_.ДатаОкончанияДействия = ActionDocument.dateEnd;
                        ActionShopDocument.ref_.НаименованиеАкции = ActionDocument.strName;
                        ActionShopDocument.ref_.ДляВсехМагазинов = true;

                        ActionShopDocument.ref_.СкидкиНаценки.Очистить();

                        // вносим скидки
                        for (int i = 0; i < ActionDocument.listDiscounts.Count(); i++)
                        {
                            dynamic new_row = ActionShopDocument.ref_.СкидкиНаценки.Добавить();
                            findDiscount = null;
                            GetDiscount(i);
                            new_row.СкидкаНаценка = findDiscount;
                            new_row.ДатаНачала = ActionDocument.dateBegin;
                            new_row.ДатаОкончания = ActionDocument.dateEnd;
                            if (new_row != null) Marshal.FinalReleaseComObject(new_row);
                            if (findDiscount != null) Marshal.FinalReleaseComObject(findDiscount);
                        }

                        if (ActionDocument.bProvodka)
                            ActionShopDocument.ref_.Записать(connectShop.РежимЗаписиДокумента.Проведение);
                        else
                            ActionShopDocument.ref_.Записать(connectShop.РежимЗаписиДокумента.ОтменаПроведения);
                    }

                    backgroundWorker1.ReportProgress(2, "Скидки успешно синхронизированы.");
                }

                //listPriceDocuments.Clear();

                if (ActionDocument != null && ActionDocument.ref_ != null) Marshal.FinalReleaseComObject(ActionDocument.ref_);
                if (ActionShopDocument != null && ActionShopDocument.ref_ != null) Marshal.FinalReleaseComObject(ActionShopDocument.ref_);

                ActionDocument = null;
                ActionShopDocument = null;
            }
            catch (Exception ex)
            {
                backgroundWorker1.ReportProgress(1, "Ошибка при синхронизации: " + ex.Message);
                return false;
            }

            return true;
        }

        private bool SyncContragent()
        {
            try
            {
                if (listSuppliers.Count() > 0 || listSuppliersShop.Count() > 0)
                {
                    backgroundWorker1.ReportProgress(0, "");
                    backgroundWorker1.ReportProgress(0, "Синхронизация контрагентов.");

                    bool bExistsCross = false;
                    for (int i = 0; i < listSuppliers.Count; i++)
                    {
                        for (int j = listSuppliersShop.Count - 1; j >= 0; j--)
                        {
                            if (listSuppliers[i].strCodeSync == listSuppliersShop[j].strCodeSync)
                            {
                                //listSuppliersShop.RemoveAt(j);
                                bExistsCross = true;
                            }
                        }
                    }

                    if (bExistsCross)
                    {
                        DialogResult res = MessageBox.Show("Некоторые контрагенты изменялись и в основной базе и в базе магазина. \nЕсли необходимо брать данные из основной базы - нажмите 'Yes' или 'Да', если из базы магазина - 'No' или 'Нет'", "Предупреждение", MessageBoxButtons.YesNo);
                        if (res == System.Windows.Forms.DialogResult.Yes)
                        {
                            for (int i = 0; i < listSuppliers.Count; i++)
                            {
                                for (int j = listSuppliersShop.Count - 1; j >= 0; j--)
                                {
                                    if (listSuppliers[i].strCodeSync == listSuppliersShop[j].strCodeSync)
                                    {
                                        listSuppliersShop.RemoveAt(j);
                                    }
                                }
                            }
                        }
                        else
                        {
                            for (int i = 0; i < listSuppliersShop.Count; i++)
                            {
                                for (int j = listSuppliers.Count - 1; j >= 0; j--)
                                {
                                    if (listSuppliers[j].strCodeSync == listSuppliersShop[i].strCodeSync)
                                    {
                                        listSuppliers.RemoveAt(j);
                                    }
                                }
                            }
                        }
                    }

                    // записываем в основную базу
                    dynamic parentVid = connect.Справочники.ВидыКонтактнойИнформации.НайтиПоНаименованию("Контактная информация справочника \"Контрагенты\"", true);

                    for (int i = 0; i < listSuppliersShop.Count; i++)
                    {
                        dynamic findSupplier = connect.Справочники.Контрагенты.НайтиПоРеквизиту("КодДляСинхронизации", listSuppliersShop[i].strCodeSync);
                        dynamic supplier = null;
                        if (findSupplier == connect.Справочники.Контрагенты.ПустаяСсылка())
                        {
                            // необходимо добавить контрагента
                            supplier = connect.Справочники.Контрагенты.СоздатьЭлемент();
                        }
                        else
                        {
                            supplier = findSupplier.ПолучитьОбъект();
                        }

                        supplier.КодДляСинхронизации = listSuppliersShop[i].strCodeSync;
                        supplier.Наименование = listSuppliersShop[i].strName;
                        supplier.НаименованиеПолное = listSuppliersShop[i].strFullName;
                        supplier.ИНН = listSuppliersShop[i].strINN;
                        supplier.Комментарий = listSuppliersShop[i].strNote;

                        supplier.Фамилия = listSuppliersShop[i].strSurname;
                        supplier.Имя = listSuppliersShop[i].strFirstname;
                        supplier.Отчество = listSuppliersShop[i].strFathername;

                        supplier.Оптовик = listSuppliersShop[i].bWholeSaler;
                        supplier.ОтказОтСМСРассылки = listSuppliersShop[i].bNoAgreeWithDelivery;
                        supplier.Игнор = listSuppliersShop[i].bIgnor;

                        supplier.ИнтересующиеМарки = listSuppliersShop[i].strBrandWishes;
                        supplier.ИнтересующиеРазмеры = listSuppliersShop[i].strSizeWishes;
                        supplier.ИнтересующиеКатегории = listSuppliersShop[i].strCategoryWishes;

                        if (listSuppliersShop[i].strWishes1 != "")
                        {
                            dynamic findWish = connect.Справочники.КонтрагентыПредпочтения.НайтиПоНаименованию(listSuppliersShop[i].strWishes1, true);
                            dynamic wish = null;
                            if (findWish == connect.Справочники.КонтрагентыПредпочтения.ПустаяСсылка())
                            {
                                // необходимо добавить тип
                                wish = connect.Справочники.КонтрагентыПредпочтения.СоздатьЭлемент();
                                wish.Наименование = listSuppliersShop[i].strWishes1;
                                wish.Записать();
                                findWish = wish.Ссылка;
                            }
                            supplier.Предпочтения1 = findWish;

                            if (findWish != null) Marshal.FinalReleaseComObject(findWish);
                        }
                        else
                            supplier.Предпочтения1 = connect.Справочники.КонтрагентыПредпочтения.ПустаяСсылка();

                        if (listSuppliersShop[i].strWishes2 != "")
                        {
                            dynamic findWish = connect.Справочники.КонтрагентыПредпочтения.НайтиПоНаименованию(listSuppliersShop[i].strWishes2, true);
                            dynamic wish = null;
                            if (findWish == connect.Справочники.КонтрагентыПредпочтения.ПустаяСсылка())
                            {
                                // необходимо добавить тип
                                wish = connect.Справочники.КонтрагентыПредпочтения.СоздатьЭлемент();
                                wish.Наименование = listSuppliersShop[i].strWishes2;
                                wish.Записать();
                                findWish = wish.Ссылка;
                            }
                            supplier.Предпочтения2 = findWish;

                            if (findWish != null) Marshal.FinalReleaseComObject(findWish);
                        }
                        else
                            supplier.Предпочтения2 = connect.Справочники.КонтрагентыПредпочтения.ПустаяСсылка();

                        if (listSuppliersShop[i].strWishes3 != "")
                        {
                            dynamic findWish = connect.Справочники.КонтрагентыПредпочтения.НайтиПоНаименованию(listSuppliersShop[i].strWishes3, true);
                            dynamic wish = null;
                            if (findWish == connect.Справочники.КонтрагентыПредпочтения.ПустаяСсылка())
                            {
                                // необходимо добавить тип
                                wish = connect.Справочники.КонтрагентыПредпочтения.СоздатьЭлемент();
                                wish.Наименование = listSuppliersShop[i].strWishes3;
                                wish.Записать();
                                findWish = wish.Ссылка;
                            }
                            supplier.Предпочтения3 = findWish;

                            if (findWish != null) Marshal.FinalReleaseComObject(findWish);
                        }
                        else
                            supplier.Предпочтения3 = connect.Справочники.КонтрагентыПредпочтения.ПустаяСсылка();

                        if (listSuppliersShop[i].Type == 0)
                            supplier.ЮрФизЛицо = connect.Перечисления.ЮрФизЛицо.ЮрЛицо;
                        else
                            supplier.ЮрФизЛицо = connect.Перечисления.ЮрФизЛицо.ФизЛицо;

                        supplier.КонтактнаяИнформация.Очистить();

                        foreach (KeyValuePair<string, string> val in listSuppliersShop[i].dictExInfo)
                        {
                            if (val.Value != "")
                            {
                                dynamic vid = connect.Справочники.ВидыКонтактнойИнформации.НайтиПоНаименованию(val.Key, true, parentVid);
                                if (vid != connect.Справочники.ВидыКонтактнойИнформации.ПустаяСсылка())
                                {
                                    dynamic new_row = supplier.КонтактнаяИнформация.Добавить();
                                    new_row.Тип = vid.Тип;
                                    new_row.Вид = vid.Ссылка;
                                    new_row.Представление = val.Value;
                                    //if (new_row != null) Marshal.FinalReleaseComObject(new_row);
                                }
                                if (vid != null) Marshal.FinalReleaseComObject(vid);
                            }
                        }

                        supplier.Записать();

                        if (supplier != null) Marshal.FinalReleaseComObject(supplier);
                        if (findSupplier != null) Marshal.FinalReleaseComObject(findSupplier);

                        backgroundWorker1.ReportProgress(0, (i + 1).ToString() + ". Контрагент " + listSuppliersShop[i].strName + " успешно синхронизирован.");

                    }
                    if (parentVid != null) Marshal.FinalReleaseComObject(parentVid);

                    // записываем в базу магазина
                    dynamic parentVidShop = connectShop.Справочники.ВидыКонтактнойИнформации.НайтиПоНаименованию("Контактная информация справочника \"Контрагенты\"", true);

                    for (int i = 0; i < listSuppliers.Count; i++)
                    {
                        dynamic findSupplier = connectShop.Справочники.Контрагенты.НайтиПоРеквизиту("КодДляСинхронизации", listSuppliers[i].strCodeSync);
                        dynamic supplier = null;
                        if (findSupplier == connectShop.Справочники.Контрагенты.ПустаяСсылка())
                        {
                            // необходимо добавить контрагента
                            supplier = connectShop.Справочники.Контрагенты.СоздатьЭлемент();
                        }
                        else
                        {
                            supplier = findSupplier.ПолучитьОбъект();
                        }

                        supplier.КодДляСинхронизации = listSuppliers[i].strCodeSync;
                        supplier.Наименование = listSuppliers[i].strName;
                        supplier.НаименованиеПолное = listSuppliers[i].strFullName;
                        supplier.ИНН = listSuppliers[i].strINN;
                        supplier.Комментарий = listSuppliers[i].strNote;

                        supplier.Фамилия = listSuppliers[i].strSurname;
                        supplier.Имя = listSuppliers[i].strFirstname;
                        supplier.Отчество = listSuppliers[i].strFathername;

                        supplier.Оптовик = listSuppliers[i].bWholeSaler;
                        supplier.ОтказОтСМСРассылки = listSuppliers[i].bNoAgreeWithDelivery;
                        supplier.Игнор = listSuppliers[i].bIgnor;

                        supplier.ИнтересующиеМарки = listSuppliers[i].strBrandWishes;
                        supplier.ИнтересующиеРазмеры = listSuppliers[i].strSizeWishes;
                        supplier.ИнтересующиеКатегории = listSuppliers[i].strCategoryWishes;

                        if (listSuppliers[i].strWishes1 != "")
                        {
                            dynamic findWish = connectShop.Справочники.КонтрагентыПредпочтения.НайтиПоНаименованию(listSuppliers[i].strWishes1, true);
                            dynamic wish = null;
                            if (findWish == connectShop.Справочники.КонтрагентыПредпочтения.ПустаяСсылка())
                            {
                                // необходимо добавить тип
                                wish = connectShop.Справочники.КонтрагентыПредпочтения.СоздатьЭлемент();
                                wish.Наименование = listSuppliers[i].strWishes1;
                                wish.Записать();
                                findWish = wish.Ссылка;
                            }
                            supplier.Предпочтения1 = findWish;

                            if (findWish != null) Marshal.FinalReleaseComObject(findWish);
                        }
                        else
                            supplier.Предпочтения1 = connectShop.Справочники.КонтрагентыПредпочтения.ПустаяСсылка();


                        if (listSuppliers[i].strWishes2 != "")
                        {
                            dynamic findWish = connectShop.Справочники.КонтрагентыПредпочтения.НайтиПоНаименованию(listSuppliers[i].strWishes2, true);
                            dynamic wish = null;
                            if (findWish == connectShop.Справочники.КонтрагентыПредпочтения.ПустаяСсылка())
                            {
                                // необходимо добавить тип
                                wish = connectShop.Справочники.КонтрагентыПредпочтения.СоздатьЭлемент();
                                wish.Наименование = listSuppliers[i].strWishes2;
                                wish.Записать();
                                findWish = wish.Ссылка;
                            }
                            supplier.Предпочтения2 = findWish;

                            if (findWish != null) Marshal.FinalReleaseComObject(findWish);
                        }
                        else
                            supplier.Предпочтения2 = connectShop.Справочники.КонтрагентыПредпочтения.ПустаяСсылка();

                        if (listSuppliers[i].strWishes3 != "")
                        {
                            dynamic findWish = connectShop.Справочники.КонтрагентыПредпочтения.НайтиПоНаименованию(listSuppliers[i].strWishes3, true);
                            dynamic wish = null;
                            if (findWish == connectShop.Справочники.КонтрагентыПредпочтения.ПустаяСсылка())
                            {
                                // необходимо добавить тип
                                wish = connectShop.Справочники.КонтрагентыПредпочтения.СоздатьЭлемент();
                                wish.Наименование = listSuppliers[i].strWishes3;
                                wish.Записать();
                                findWish = wish.Ссылка;
                            }
                            supplier.Предпочтения3 = findWish;

                            if (findWish != null) Marshal.FinalReleaseComObject(findWish);
                        }
                        else
                            supplier.Предпочтения3 = connectShop.Справочники.КонтрагентыПредпочтения.ПустаяСсылка();


                        if (listSuppliers[i].Type == 0)
                            supplier.ЮрФизЛицо = connectShop.Перечисления.ЮрФизЛицо.ЮрЛицо;
                        else
                            supplier.ЮрФизЛицо = connectShop.Перечисления.ЮрФизЛицо.ФизЛицо;

                        supplier.КонтактнаяИнформация.Очистить();

                        foreach (KeyValuePair<string, string> val in listSuppliers[i].dictExInfo)
                        {
                            if (val.Value != "")
                            {
                                dynamic vid = connectShop.Справочники.ВидыКонтактнойИнформации.НайтиПоНаименованию(val.Key, true, parentVidShop);
                                if (vid != connectShop.Справочники.ВидыКонтактнойИнформации.ПустаяСсылка())
                                {
                                    dynamic new_row = supplier.КонтактнаяИнформация.Добавить();
                                    new_row.Тип = vid.Тип;
                                    new_row.Вид = vid.Ссылка;
                                    new_row.Представление = val.Value;
                                    if (new_row != null) Marshal.FinalReleaseComObject(new_row);
                                }
                                if (vid != null) Marshal.FinalReleaseComObject(vid);
                            }
                        }

                        supplier.Записать();

                        if (supplier != null) Marshal.FinalReleaseComObject(supplier);
                        if (findSupplier != null) Marshal.FinalReleaseComObject(findSupplier);

                        backgroundWorker1.ReportProgress(0, (i + 1).ToString() + ". Контрагент " + listSuppliers[i].strName + " успешно синхронизирован.");
                    }

                    if (parentVidShop != null) Marshal.FinalReleaseComObject(parentVidShop);

                    listSuppliers.Clear();
                    listSuppliersShop.Clear();

                    backgroundWorker1.ReportProgress(2, "Контрагенты успешно синхронизированы.");
                }
            }
            catch (Exception ex)
            {
                backgroundWorker1.ReportProgress(1, "Ошибка при синхронизации: " + ex.Message);
                return false;
            }

            return true;
        }

        private void GetDiscount(int index)
        {
            findDiscount = connectShop.Справочники.СкидкиНаценки.НайтиПоНаименованию(ActionDocument.listDiscounts[index].strName, true);
            dynamic discount = null;
            if (findDiscount != connectShop.Справочники.СкидкиНаценки.ПустаяСсылка())
            {
                if (findDiscount.ЗначениеСкидкиНаценки != ActionDocument.listDiscounts[index].fProcent)
                {
                    discount = findDiscount.ПолучитьОбъект();
                    discount.ЗначениеСкидкиНаценки = ActionDocument.listDiscounts[index].fProcent;
                    discount.Записать();
                }
            }
            else
            {
                dynamic findGroup = connectShop.Справочники.СкидкиНаценки.НайтиПоНаименованию("Группа", true);
                dynamic group = null;
                if (findGroup == connectShop.Справочники.СкидкиНаценки.ПустаяСсылка())
                {
                    group = connectShop.Справочники.СкидкиНаценки.СоздатьГруппу();
                    group.Наименование = "Группа";
                    group.ВариантСовместногоПрименения = connectShop.Перечисления.ВариантыСовместногоПримененияСкидокНаценок.Вытеснение;
                    group.Записать();

                    findGroup = group.Ссылка;
                }

                discount = connectShop.Справочники.СкидкиНаценки.СоздатьЭлемент();
                discount.Наименование = ActionDocument.listDiscounts[index].strName;
                discount.ЗначениеСкидкиНаценки = ActionDocument.listDiscounts[index].fProcent;
                discount.СпособПредоставления = connectShop.Перечисления.СпособыПредоставленияСкидокНаценок.Процент;
                discount.СтатусДействия = connectShop.Перечисления.СтатусыДействияСкидок.Действует;
                discount.ОбластьПредоставления = connectShop.Перечисления.ВариантыОбластейОграниченияСкидокНаценок.ВДокументе;
                discount.Родитель = findGroup;

                dynamic new_row = discount.УсловияПредоставления.Добавить();
                findDiscountCond = null;
                GetDiscountCond(index); // Дисконтная карта: опт -10%
                new_row.УсловиеПредоставления = findDiscountCond;
                if (new_row != null) Marshal.FinalReleaseComObject(new_row);
                if (findDiscountCond != null) Marshal.FinalReleaseComObject(findDiscountCond);

                discount.Записать();

                findDiscount = discount.Ссылка;

                //if (group != null) Marshal.FinalReleaseComObject(group);
                if (findGroup != null) Marshal.FinalReleaseComObject(findGroup);
            }

            if (discount != null) Marshal.FinalReleaseComObject(discount);
            //return findDiscount;
        }

        private void GetDiscountCond(int index)
        {
            dynamic dc = ActionDocument.ref_.СкидкиНаценки.Выгрузить().Получить(index).СкидкаНаценка.УсловияПредоставления.Выгрузить().Получить(0).УсловиеПредоставления;
            string strDiscountCondName = dc.Наименование;

            findDiscountCond = connectShop.Справочники.УсловияПредоставленияСкидокНаценок.НайтиПоНаименованию(strDiscountCondName, true);
            dynamic discountCond = null;

            if (findDiscountCond == connectShop.Справочники.УсловияПредоставленияСкидокНаценок.ПустаяСсылка())
            {
                if (connect.String(dc.УсловиеПредоставления) == connect.String(connect.Перечисления.УсловияПредоставленияСкидокНаценок.ПоТипуПолучателя))
                {
                    dynamic ic = dc.Получатели.Выгрузить().Получить(0);
                    string strDiscountCardName = ic.Получатель.Наименование;
                    string strDiscountCardCode = ic.Получатель.КодКарты;

                    discountCond = connectShop.Справочники.УсловияПредоставленияСкидокНаценок.СоздатьЭлемент();
                    discountCond.Наименование = strDiscountCondName;
                    discountCond.УсловиеПредоставления = connectShop.Перечисления.УсловияПредоставленияСкидокНаценок.ПоТипуПолучателя;
                    dynamic new_row = discountCond.Получатели.Добавить();
                    findDiscountCard = null;
                    GetDiscountCard(strDiscountCardName, strDiscountCardCode); // опт -10%
                    new_row.Получатель = findDiscountCard; 
                    if (new_row != null) Marshal.FinalReleaseComObject(new_row);
                    if (findDiscountCard != null) Marshal.FinalReleaseComObject(findDiscountCard);

                    discountCond.Записать();

                    findDiscountCond = discountCond.Ссылка;

                    if (ic != null) Marshal.FinalReleaseComObject(ic);
                }
                else
                {
                    // иначе это наша доставка
                    MessageBox.Show("Необходимо вручную установить нулевую скидку по номенклатуре \"доставка\" с наименованием - \"" + strDiscountCondName + "\"");

                    /*discountCond = connectShop.Справочники.УсловияПредоставленияСкидокНаценок.СоздатьЭлемент();
                    discountCond.Наименование = strDiscountCondName;
                    discountCond.УсловиеПредоставления = connectShop.Перечисления.УсловияПредоставленияСкидокНаценок.ЗаКомплектПокупки;
                    dynamic new_row = discountCond.КомплектПокупки.Добавить();

                    dynamic tovar = connectShop.Справочники.Номенклатура.НайтиПоНаименованию("доставка", true);
                    new_row.Номенклатура = tovar;
                    new_row.Характеристика = connectShop.Справочники.ХарактеристикиНоменклатуры.НайтиПоНаименованию("<не указан>, <не указан>", true, null,tovar);
                    new_row.КоличествоУпаковок = 1;
                    new_row.Количество = 1;
                    if (new_row != null) Marshal.FinalReleaseComObject(new_row);
                    if (tovar != null) Marshal.FinalReleaseComObject(tovar);

                    discountCond.Записать();

                    findDiscountCond = discountCond.Ссылка;*/

                }
            }

            if (discountCond != null) Marshal.FinalReleaseComObject(discountCond);
            if (dc != null) Marshal.FinalReleaseComObject(dc);
        }

        private void GetDiscountCard(string strName, string strCode)
        {
            findDiscountCard = connectShop.Справочники.ИнформационныеКарты.НайтиПоНаименованию(strName, true);
            dynamic discountCard = null;
            if (findDiscountCard == connectShop.Справочники.ИнформационныеКарты.ПустаяСсылка())
            {
                discountCard = connectShop.Справочники.ИнформационныеКарты.СоздатьЭлемент();
                discountCard.Наименование = strName;
                discountCard.КодКарты = strCode;
                discountCard.ВидКарты = connectShop.Перечисления.ВидыИнформационныхКарт.Магнитная;
                discountCard.ТипКарты = connectShop.Перечисления.ТипыИнформационныхКарт.Дисконтная;

                // поиск вида
                // ------------------------
                dynamic findType = connectShop.Справочники.ВидыДисконтныхКарт.НайтиПоНаименованию("Вид дисконтной карты", true);
                dynamic type = null;
                if (findType == connectShop.Справочники.ВидыДисконтныхКарт.ПустаяСсылка())
                {
                    // необходимо добавить тип
                    type = connectShop.Справочники.ВидыДисконтныхКарт.СоздатьЭлемент();
                    type.Наименование = "Вид дисконтной карты";
                    type.Записать();
                    findType = type.Ссылка;
                }
                // ------------------------

                discountCard.ВидДисконтнойКарты = findType;
                discountCard.Записать();

                findDiscountCard = discountCard.Ссылка;

                if (findType != null) Marshal.FinalReleaseComObject(findType);
                if (type != null) Marshal.FinalReleaseComObject(type);
            }

            if (discountCard != null) Marshal.FinalReleaseComObject(discountCard);
            //return findDiscountCard;
        }

        private bool SyncInvoiceIn()
        {
            try
            {
                // сначала получаем накладные из магазина, потом скидываем в магазин
                // если накладная менялась и в основной базе и в магазине, то главнее накладная основной базы
                if (listInvoiceInDocumentsShop.Count() > 0)
                {
                    backgroundWorker1.ReportProgress(0, "");
                    backgroundWorker1.ReportProgress(0, "Синхронизация документов получения товаров");
                }

                // передаем данные по приходу товаров
                for (int i = 0; i < listInvoiceInDocumentsShop.Count(); i++)
                {

                    // если не находим документа с таким номером и датой, просто добавляем его и проводим
                    // если находим, тогда добавляем только позиции

                    dynamic findDoc = connect.Документы.ПоступлениеТоваров.НайтиПоРеквизиту("КодДляСинхронизации", listInvoiceInDocumentsShop[i].strCodeSync);
                    dynamic doc = null;
                    if (findDoc == connect.Документы.ПоступлениеТоваров.ПустаяСсылка())
                    {
                        // если новый и не проведен, тогда не переносим
                        if (!listInvoiceInDocumentsShop[i].bProvodka)
                        {
                            if (findDoc != null) Marshal.FinalReleaseComObject(findDoc);
                            if (doc != null) Marshal.FinalReleaseComObject(doc);

                            backgroundWorker1.ReportProgress(0, (i + 1).ToString() + ". Документ поступления №" + listInvoiceInDocumentsShop[i].strCode + " от " + listInvoiceInDocumentsShop[i].date.ToShortDateString() + " пропущен, т.к. он не проведен.");

                            continue;
                        }
                        // необходимо добавить номенклатуру
                        doc = connect.Документы.ПоступлениеТоваров.СоздатьДокумент();
                    }
                    else
                    {
                        /*if (MessageBox.Show("Документ поступления №" + listInvoiceInDocumentsShop[i].strCode + " от " + listInvoiceInDocumentsShop[i].date.ToShortDateString() + " уже присутствует в основной базе. Изменить его?", "Предупреждение", MessageBoxButtons.OKCancel) == System.Windows.Forms.DialogResult.Cancel)
                        {
                            continue;
                        }*/

                        doc = findDoc.ПолучитьОбъект();
                        /*doc.Удалить();

                        if (findDoc != null) Marshal.FinalReleaseComObject(findDoc);
                        if (doc != null) Marshal.FinalReleaseComObject(doc);
                        findDoc = null;

                        doc = connect.Документы.ПоступлениеТоваров.СоздатьДокумент();

                        // необходимо изменить номенклатуру
                        //doc = findDoc.ПолучитьОбъект();*/
                        //bProvodka = doc.Проведен;
                        //bDeleted = doc.ПометкаУдаления;
                    }

                    dynamic findUser = connect.Справочники.Пользователи.НайтиПоНаименованию(listInvoiceInDocumentsShop[i].strUser, true);
                    if (findUser == connect.Справочники.Пользователи.ПустаяСсылка())
                    {
                        dynamic user = connect.Справочники.Пользователи.СоздатьЭлемент();
                        user.Наименование = listInvoiceInDocumentsShop[i].strUser;
                        user.Записать();

                        findUser = user.Ссылка;

                        if (user != null) Marshal.FinalReleaseComObject(user);
                    }

                    // поиск контрагента
                    // ------------------------
                    dynamic findContragent = connect.Справочники.Контрагенты.НайтиПоРеквизиту("КодДляСинхронизации", listInvoiceInDocumentsShop[i].strSupplierCodeSync);
                    dynamic contragent = null;
                    if (findContragent == connect.Справочники.Контрагенты.ПустаяСсылка())
                    {
                        // необходимо добавить контрагента

                        // получаем
                        dynamic objects = connectShop.Справочники.Контрагенты.НайтиПоРеквизиту("КодДляСинхронизации", listInvoiceInDocumentsShop[i].strSupplierCodeSync);

                        SupplierInfo si;
                        if (objects != connectShop.Справочники.Контрагенты.ПустаяССылка())
                        {
                            si = new SupplierInfo();
                            si.strCodeSync = objects.КодДляСинхронизации;
                            si.strName = objects.Наименование;
                            si.strFullName = objects.НаименованиеПолное;
                            si.strINN = objects.ИНН;
                            si.strNote = objects.Комментарий;

                            si.strSurname = objects.Ссылка.Фамилия;
                            si.strFirstname = objects.Ссылка.Имя;
                            si.strFathername = objects.Ссылка.Отчество;

                            si.bWholeSaler = objects.Ссылка.Оптовик;
                            si.bNoAgreeWithDelivery = objects.Ссылка.ОтказОтСМСРассылки;
                            si.bIgnor = objects.Ссылка.Игнор;

                            si.strBrandWishes = objects.Ссылка.ИнтересующиеМарки;
                            si.strSizeWishes = objects.Ссылка.ИнтересующиеРазмеры;
                            si.strCategoryWishes = objects.Ссылка.ИнтересующиеКатегории;

                            si.strWishes1 = "";
                            si.strWishes2 = "";
                            si.strWishes3 = "";

                            if (objects.Ссылка.Предпочтения1 != connectShop.Справочники.КонтрагентыПредпочтения.ПустаяСсылка())
                                si.strWishes1 = objects.Ссылка.Предпочтения1.Наименование;
                            if (objects.Ссылка.Предпочтения2 != connectShop.Справочники.КонтрагентыПредпочтения.ПустаяСсылка())
                                si.strWishes2 = objects.Ссылка.Предпочтения2.Наименование;
                            if (objects.Ссылка.Предпочтения3 != connectShop.Справочники.КонтрагентыПредпочтения.ПустаяСсылка())
                                si.strWishes3 = objects.Ссылка.Предпочтения3.Наименование;

                            si.dateChange = objects.ДатаИзменения;
                            dynamic val_ = objects.ЮрФизЛицо;
                            if (connectShop.String(val_) == connectShop.String(connectShop.Перечисления.ЮрФизЛицо.ЮрЛицо))
                                si.Type = 0;
                            else
                                si.Type = 1;

                            dynamic ex_info = objects.КонтактнаяИнформация.Выгрузить();
                            int count = ex_info.Количество();

                            for (int ii = 0; ii < count; ii++)
                            {
                                dynamic row = ex_info.Получить(ii);
                                si.dictExInfo[row.Вид.Наименование] = row.Представление;
                                Marshal.FinalReleaseComObject(row);
                            }
                            Marshal.FinalReleaseComObject(val_);

                            // добавляем
                            dynamic supplier = null;
                            supplier = connect.Справочники.Контрагенты.СоздатьЭлемент();

                            supplier.КодДляСинхронизации = si.strCodeSync;
                            supplier.Наименование = si.strName;
                            supplier.НаименованиеПолное = si.strFullName;
                            supplier.ИНН = si.strINN;
                            supplier.Комментарий = si.strNote;

                            supplier.Фамилия = si.strSurname;
                            supplier.Имя = si.strFirstname;
                            supplier.Отчество = si.strFathername;

                            supplier.Оптовик = si.bWholeSaler;
                            supplier.ОтказОтСМСРассылки = si.bNoAgreeWithDelivery;
                            supplier.Игнор = si.bIgnor;

                            supplier.ИнтересующиеМарки = si.strBrandWishes;
                            supplier.ИнтересующиеРазмеры = si.strSizeWishes;
                            supplier.ИнтересующиеКатегории = si.strCategoryWishes;

                            if (supplier.strWishes1 != "")
                            {
                                dynamic findWish = connect.Справочники.КонтрагентыПредпочтения.НайтиПоНаименованию(supplier.strWishes1, true);
                                dynamic wish = null;
                                if (findWish == connect.Справочники.КонтрагентыПредпочтения.ПустаяСсылка())
                                {
                                    // необходимо добавить тип
                                    wish = connect.Справочники.КонтрагентыПредпочтения.СоздатьЭлемент();
                                    wish.Наименование = supplier.strWishes1;
                                    wish.Записать();
                                    findWish = wish.Ссылка;
                                }
                                supplier.Предпочтения1 = findWish;

                                if (findWish != null) Marshal.FinalReleaseComObject(findWish);
                            }
                            else
                                supplier.Предпочтения1 = connect.Справочники.КонтрагентыПредпочтения.ПустаяСсылка();


                            if (supplier.strWishes2 != "")
                            {
                                dynamic findWish = connect.Справочники.КонтрагентыПредпочтения.НайтиПоНаименованию(supplier.strWishes2, true);
                                dynamic wish = null;
                                if (findWish == connect.Справочники.КонтрагентыПредпочтения.ПустаяСсылка())
                                {
                                    // необходимо добавить тип
                                    wish = connect.Справочники.КонтрагентыПредпочтения.СоздатьЭлемент();
                                    wish.Наименование = supplier.strWishes2;
                                    wish.Записать();
                                    findWish = wish.Ссылка;
                                }
                                supplier.Предпочтения2 = findWish;

                                if (findWish != null) Marshal.FinalReleaseComObject(findWish);
                            }
                            else
                                supplier.Предпочтения2 = connect.Справочники.КонтрагентыПредпочтения.ПустаяСсылка();


                            if (supplier.strWishes3 != "")
                            {
                                dynamic findWish = connect.Справочники.КонтрагентыПредпочтения.НайтиПоНаименованию(supplier.strWishes3, true);
                                dynamic wish = null;
                                if (findWish == connect.Справочники.КонтрагентыПредпочтения.ПустаяСсылка())
                                {
                                    // необходимо добавить тип
                                    wish = connect.Справочники.КонтрагентыПредпочтения.СоздатьЭлемент();
                                    wish.Наименование = supplier.strWishes3;
                                    wish.Записать();
                                    findWish = wish.Ссылка;
                                }
                                supplier.Предпочтения3 = findWish;

                                if (findWish != null) Marshal.FinalReleaseComObject(findWish);
                            }
                            else
                                supplier.Предпочтения3 = connect.Справочники.КонтрагентыПредпочтения.ПустаяСсылка();


                            if (si.Type == 0)
                                supplier.ЮрФизЛицо = connect.Перечисления.ЮрФизЛицо.ЮрЛицо;
                            else
                                supplier.ЮрФизЛицо = connect.Перечисления.ЮрФизЛицо.ФизЛицо;

                            supplier.КонтактнаяИнформация.Очистить();

                            foreach (KeyValuePair<string, string> val in si.dictExInfo)
                            {
                                if (val.Value != "")
                                {
                                    dynamic parent = connect.Справочники.ВидыКонтактнойИнформации.НайтиПоНаименованию("Контактная информация справочника \"Контрагенты\"", true);
                                    dynamic vid = connect.Справочники.ВидыКонтактнойИнформации.НайтиПоНаименованию(val.Key, true, parent);
                                    if (vid != connect.Справочники.ВидыКонтактнойИнформации.ПустаяСсылка())
                                    {
                                        dynamic new_row = supplier.КонтактнаяИнформация.Добавить();
                                        new_row.Тип = vid.Тип;
                                        new_row.Вид = vid.Ссылка;
                                        new_row.Представление = val.Value;
                                        //if (new_row != null) Marshal.FinalReleaseComObject(new_row);
                                    }
                                    if (vid != null) Marshal.FinalReleaseComObject(vid);
                                    if (parent != null) Marshal.FinalReleaseComObject(parent);
                                }
                            }

                            supplier.Записать();

                            findContragent = supplier.Ссылка;

                            if (supplier != null) Marshal.FinalReleaseComObject(supplier);

                            if (objects != null) Marshal.FinalReleaseComObject(objects);
                        }
                        else
                        {
                            // необходимо добавить контрагента
                            if (objects != null) Marshal.FinalReleaseComObject(objects);
                            if (findContragent != null) Marshal.FinalReleaseComObject(findContragent);
                            if (contragent != null) Marshal.FinalReleaseComObject(contragent);
                            if (findUser != null) Marshal.FinalReleaseComObject(findUser);
                            if (findDoc != null) Marshal.FinalReleaseComObject(findDoc);
                            if (doc != null) Marshal.FinalReleaseComObject(doc);
                            backgroundWorker1.ReportProgress(1, (i + 1).ToString() + ". Документ поступления №" + listInvoiceInDocumentsShop[i].strCode + " от " + listInvoiceInDocumentsShop[i].date.ToShortDateString() + " пропущен, т.к. не удалось найти контрагента.");
                            continue;
                        }
                    }
                    // ------------------------

                    doc.КодДляСинхронизации = listInvoiceInDocumentsShop[i].strCodeSync;
                    //doc.Номер = listInvoiceInDocumentsShop[i].strCode;
                    doc.Отчет = listInvoiceInDocumentsShop[i].bReport;
                    doc.Дата = listInvoiceInDocumentsShop[i].date;
                    doc.Организация = ref_company.Ссылка;
                    doc.Магазин = ref_shop.Ссылка;
                    doc.Склад = ref_shop.СкладПоступления.Ссылка;
                    doc.Контрагент = findContragent;
                    if (listInvoiceInDocumentsShop[i].bDeleted) doc.ПометкаУдаления = true;
                    doc.УчитыватьНДС = true;
                    doc.ЦенаВключаетНДС = true;
                    doc.Ответственный = findUser;
                    doc.Комментарий = listInvoiceInDocumentsShop[i].strNote;
                    doc.Товары.Очистить();

                    // добавление товаров
                    bool bError = false;
                    //double price = 0;
                    for (int j = 0; j < listInvoiceInDocumentsShop[i].listTovars.Count(); j++)
                    {
                        dynamic findTovar = connect.Справочники.Номенклатура.НайтиПоРеквизиту("КодДляСинхронизации", listInvoiceInDocumentsShop[i].listTovars[j].strTovarSyncCode);
                        if (findTovar == connect.Справочники.Номенклатура.ПустаяСсылка())
                        {
                            backgroundWorker1.ReportProgress(1, (i + 1).ToString() + ". Документ поступления №" + listInvoiceInDocumentsShop[i].strCode + " от " + listInvoiceInDocumentsShop[i].date.ToShortDateString() + " пропущен, т.к. не удалось найти номенклатуру с кодом " + listInvoiceInDocumentsShop[i].listTovars[j].strTovarSyncCode);
                            if (findTovar != null) Marshal.FinalReleaseComObject(findTovar);
                            bError = true;
                            break;
                        }
                        dynamic findAttrib = connect.Справочники.ХарактеристикиНоменклатуры.НайтиПоРеквизиту("КодДляСинхронизации", listInvoiceInDocumentsShop[i].listTovars[j].strAttribSyncCode);
                        if (findAttrib == connect.Справочники.ХарактеристикиНоменклатуры.ПустаяСсылка())
                        {
                            backgroundWorker1.ReportProgress(1, (i + 1).ToString() + ". Документ поступления №" + listInvoiceInDocumentsShop[i].strCode + " от " + listInvoiceInDocumentsShop[i].date.ToShortDateString() + " пропущен, т.к. не удалось найти характеристику номенклатуры с кодом " + listInvoiceInDocumentsShop[i].listTovars[j].strAttribSyncCode);
                            if (findTovar != null) Marshal.FinalReleaseComObject(findTovar);
                            if (findAttrib != null) Marshal.FinalReleaseComObject(findTovar);
                            bError = true;
                            break;
                        }

                        dynamic new_row = doc.Товары.Добавить();
                        new_row.Номенклатура = findTovar;
                        new_row.СтавкаНДС = connect.Перечисления.СтавкиНДС.БезНДС;
                        new_row.Количество = listInvoiceInDocumentsShop[i].listTovars[j].iCount;
                        new_row.КоличествоУпаковок = listInvoiceInDocumentsShop[i].listTovars[j].iCount;
                        new_row.Цена = listInvoiceInDocumentsShop[i].listTovars[j].price;
                        new_row.Сумма = new_row.Цена * new_row.Количество;
                        new_row.Характеристика = findAttrib;

                        //price = listInvoiceInDocumentsShop[i].listTovars[j].price * listInvoiceInDocumentsShop[i].listTovars[j].iCount;

                        if (new_row != null) Marshal.FinalReleaseComObject(new_row);
                        if (findTovar != null) Marshal.FinalReleaseComObject(findTovar);
                        if (findAttrib != null) Marshal.FinalReleaseComObject(findAttrib);
                    }

                    if (bError)
                    {
                        if (findContragent != null) Marshal.FinalReleaseComObject(findContragent);
                        if (contragent != null) Marshal.FinalReleaseComObject(contragent);
                        if (findUser != null) Marshal.FinalReleaseComObject(findUser);
                        if (findDoc != null) Marshal.FinalReleaseComObject(findDoc);
                        if (doc != null) Marshal.FinalReleaseComObject(doc);

                        continue;
                    }

                    //doc.СуммаДокумента = listInvoiceInDocumentsShop[i].summa;

                    if (listInvoiceInDocumentsShop[i].bProvodka)
                        doc.Записать(connect.РежимЗаписиДокумента.Проведение);
                    else
                        doc.Записать(connect.РежимЗаписиДокумента.ОтменаПроведения);

                    if (findContragent != null) Marshal.FinalReleaseComObject(findContragent);
                    if (contragent != null) Marshal.FinalReleaseComObject(contragent);
                    if (findUser != null) Marshal.FinalReleaseComObject(findUser);
                    if (findDoc != null) Marshal.FinalReleaseComObject(findDoc);
                    if (doc != null) Marshal.FinalReleaseComObject(doc);

                    backgroundWorker1.ReportProgress(0, (i + 1).ToString() + ". Документ поступления №" + listInvoiceInDocumentsShop[i].strCode + " от " + listInvoiceInDocumentsShop[i].date.ToShortDateString() + " успешно синхронизирован.");
                }

                if (listInvoiceInDocumentsShop.Count() > 0)
                    backgroundWorker1.ReportProgress(2, "Документы получения товаров успешно синхронизированы.");


                //listInvoiceInDocuments.Clear();
                listInvoiceInDocumentsShop.Clear();
            }
            catch (Exception ex)
            {
                backgroundWorker1.ReportProgress(1, "Ошибка при синхронизации документов получения товаров: " + ex.Message);
                return false;
            }

            return true;
        }

        private bool SyncInvoiceOut()
        {
            try
            {
                if (listInvoiceOutDocumentsShop.Count() > 0)
                {
                    backgroundWorker1.ReportProgress(0, "");
                    backgroundWorker1.ReportProgress(0, "Синхронизация документов реализации товаров");
                }

                // передаем данные по приходу товаров
                for (int i = 0; i < listInvoiceOutDocumentsShop.Count(); i++)
                {

                    //try
                    {
                        // если не находим документа с таким номером и датой, просто добавляем его и проводим
                        // если находим, тогда добавляем только позиции
                        backgroundWorker1.ReportProgress(0, (i + 1).ToString() + ". Обработка документа реализации №" + listInvoiceOutDocumentsShop[i].strCode + " от " + listInvoiceOutDocumentsShop[i].date.ToShortDateString() + ".");

                        dynamic findDoc = connect.Документы.РеализацияТоваров.НайтиПоРеквизиту("КодДляСинхронизации", listInvoiceOutDocumentsShop[i].strCodeSync);
                        dynamic doc = null;
                        if (findDoc == connect.Документы.РеализацияТоваров.ПустаяСсылка())
                        {
                            // если новый и не проведен, тогда не переносим
                            if (!listInvoiceOutDocumentsShop[i].bProvodka)
                            {
                                if (findDoc != null) Marshal.FinalReleaseComObject(findDoc);
                                if (doc != null) Marshal.FinalReleaseComObject(doc);
                                backgroundWorker1.ReportProgress(0, (i + 1).ToString() + ". Документ реализации №" + listInvoiceOutDocumentsShop[i].strCode + " от " + listInvoiceOutDocumentsShop[i].date.ToShortDateString() + " пропущен, т.к. он не проведен.");
                                continue;
                            }
                            // необходимо добавить номенклатуру
                            doc = connect.Документы.РеализацияТоваров.СоздатьДокумент();
                        }
                        else
                        {
                            // необходимо изменить документ, но он почему то не изменяется, поэтому просто удаляем этот документ и вносим новый
                            /*if (MessageBox.Show("Документ реализации №" + listInvoiceOutDocumentsShop[i].strCode + " от " + listInvoiceOutDocumentsShop[i].date.ToShortDateString() + " уже присутствует в основной базе. Изменить его?", "Предупреждение", MessageBoxButtons.OKCancel) == System.Windows.Forms.DialogResult.Cancel)
                            {
                                continue;
                            }*/

                            doc = findDoc.ПолучитьОбъект();
                            /*doc.Удалить();

                            if (findDoc != null) Marshal.FinalReleaseComObject(findDoc);
                            if (doc != null) Marshal.FinalReleaseComObject(doc);
                            findDoc = null;

                            doc = connect.Документы.РеализацияТоваров.СоздатьДокумент();*/
                        }

                        dynamic findUser = connect.Справочники.Пользователи.НайтиПоНаименованию(listInvoiceOutDocumentsShop[i].strUser, true);
                        if (findUser == connect.Справочники.Пользователи.ПустаяСсылка())
                        {
                            dynamic user = connect.Справочники.Пользователи.СоздатьЭлемент();
                            user.Наименование = listInvoiceOutDocumentsShop[i].strUser;
                            user.Записать();

                            findUser = user.Ссылка;

                            if (user != null) Marshal.FinalReleaseComObject(user);
                        }

                        dynamic findContragent = connect.Справочники.Контрагенты.НайтиПоРеквизиту("КодДляСинхронизации", listInvoiceOutDocumentsShop[i].strSupplierCodeSync);
                        dynamic contragent = null;
                        if (findContragent == connect.Справочники.Контрагенты.ПустаяСсылка())
                        {
                            // необходимо добавить контрагента

                            // получаем
                            dynamic objects = connectShop.Справочники.Контрагенты.НайтиПоРеквизиту("КодДляСинхронизации", listInvoiceOutDocumentsShop[i].strSupplierCodeSync);

                            SupplierInfo si;
                            if (objects != connectShop.Справочники.Контрагенты.ПустаяССылка())
                            {
                                si = new SupplierInfo();
                                si.strCodeSync = objects.КодДляСинхронизации;
                                si.strName = objects.Наименование;
                                si.strFullName = objects.НаименованиеПолное;
                                si.strINN = objects.ИНН;
                                si.strNote = objects.Комментарий;

                                si.strSurname = objects.Ссылка.Фамилия;
                                si.strFirstname = objects.Ссылка.Имя;
                                si.strFathername = objects.Ссылка.Отчество;

                                si.bWholeSaler = objects.Ссылка.Оптовик;
                                si.bNoAgreeWithDelivery = objects.Ссылка.ОтказОтСМСРассылки;
                                si.bIgnor = objects.Ссылка.Игнор;

                                si.strBrandWishes = objects.Ссылка.ИнтересующиеМарки;
                                si.strSizeWishes = objects.Ссылка.ИнтересующиеРазмеры;
                                si.strCategoryWishes = objects.Ссылка.ИнтересующиеКатегории;

                                si.strWishes1 = "";
                                si.strWishes2 = "";
                                si.strWishes3 = "";

                                if (objects.Ссылка.Предпочтения1 != connectShop.Справочники.КонтрагентыПредпочтения.ПустаяСсылка())
                                    si.strWishes1 = objects.Ссылка.Предпочтения1.Наименование;
                                if (objects.Ссылка.Предпочтения2 != connectShop.Справочники.КонтрагентыПредпочтения.ПустаяСсылка())
                                    si.strWishes2 = objects.Ссылка.Предпочтения2.Наименование;
                                if (objects.Ссылка.Предпочтения3 != connectShop.Справочники.КонтрагентыПредпочтения.ПустаяСсылка())
                                    si.strWishes3 = objects.Ссылка.Предпочтения3.Наименование;

                                si.dateChange = objects.ДатаИзменения;
                                dynamic val_ = objects.ЮрФизЛицо;
                                if (connectShop.String(val_) == connectShop.String(connectShop.Перечисления.ЮрФизЛицо.ЮрЛицо))
                                    si.Type = 0;
                                else
                                    si.Type = 1;

                                dynamic ex_info = objects.КонтактнаяИнформация.Выгрузить();
                                int count = ex_info.Количество();

                                for (int ii = 0; ii < count; ii++)
                                {
                                    dynamic row = ex_info.Получить(ii);
                                    si.dictExInfo[row.Вид.Наименование] = row.Представление;
                                    Marshal.FinalReleaseComObject(row);
                                }
                                Marshal.FinalReleaseComObject(val_);

                                // добавляем
                                dynamic supplier = null;
                                supplier = connect.Справочники.Контрагенты.СоздатьЭлемент();

                                supplier.КодДляСинхронизации = si.strCodeSync;
                                supplier.Наименование = si.strName;
                                supplier.НаименованиеПолное = si.strFullName;
                                supplier.ИНН = si.strINN;
                                supplier.Комментарий = si.strNote;

                                supplier.Фамилия = si.strSurname;
                                supplier.Имя = si.strFirstname;
                                supplier.Отчество = si.strFathername;

                                supplier.Оптовик = si.bWholeSaler;
                                supplier.ОтказОтСМСРассылки = si.bNoAgreeWithDelivery;
                                supplier.Игнор = si.bIgnor;

                                supplier.ИнтересующиеМарки = si.strBrandWishes;
                                supplier.ИнтересующиеРазмеры = si.strSizeWishes;
                                supplier.ИнтересующиеКатегории = si.strCategoryWishes;

                                if (supplier.strWishes1 != "")
                                {
                                    dynamic findWish = connect.Справочники.КонтрагентыПредпочтения.НайтиПоНаименованию(supplier.strWishes1, true);
                                    dynamic wish = null;
                                    if (findWish == connect.Справочники.КонтрагентыПредпочтения.ПустаяСсылка())
                                    {
                                        // необходимо добавить тип
                                        wish = connect.Справочники.КонтрагентыПредпочтения.СоздатьЭлемент();
                                        wish.Наименование = supplier.strWishes1;
                                        wish.Записать();
                                        findWish = wish.Ссылка;
                                    }
                                    supplier.Предпочтения1 = findWish;

                                    if (findWish != null) Marshal.FinalReleaseComObject(findWish);
                                }
                                else
                                    supplier.Предпочтения1 = connect.Справочники.КонтрагентыПредпочтения.ПустаяСсылка();

                                if (supplier.strWishes2 != "")
                                {
                                    dynamic findWish = connect.Справочники.КонтрагентыПредпочтения.НайтиПоНаименованию(supplier.strWishes2, true);
                                    dynamic wish = null;
                                    if (findWish == connect.Справочники.КонтрагентыПредпочтения.ПустаяСсылка())
                                    {
                                        // необходимо добавить тип
                                        wish = connect.Справочники.КонтрагентыПредпочтения.СоздатьЭлемент();
                                        wish.Наименование = supplier.strWishes2;
                                        wish.Записать();
                                        findWish = wish.Ссылка;
                                    }
                                    supplier.Предпочтения2 = findWish;

                                    if (findWish != null) Marshal.FinalReleaseComObject(findWish);
                                }
                                else
                                    supplier.Предпочтения2 = connect.Справочники.КонтрагентыПредпочтения.ПустаяСсылка();

                                if (supplier.strWishes3 != "")
                                {
                                    dynamic findWish = connect.Справочники.КонтрагентыПредпочтения.НайтиПоНаименованию(supplier.strWishes3, true);
                                    dynamic wish = null;
                                    if (findWish == connect.Справочники.КонтрагентыПредпочтения.ПустаяСсылка())
                                    {
                                        // необходимо добавить тип
                                        wish = connect.Справочники.КонтрагентыПредпочтения.СоздатьЭлемент();
                                        wish.Наименование = supplier.strWishes3;
                                        wish.Записать();
                                        findWish = wish.Ссылка;
                                    }
                                    supplier.Предпочтения3 = findWish;

                                    if (findWish != null) Marshal.FinalReleaseComObject(findWish);
                                }
                                else
                                    supplier.Предпочтения3 = connect.Справочники.КонтрагентыПредпочтения.ПустаяСсылка();

                                if (si.Type == 0)
                                    supplier.ЮрФизЛицо = connect.Перечисления.ЮрФизЛицо.ЮрЛицо;
                                else
                                    supplier.ЮрФизЛицо = connect.Перечисления.ЮрФизЛицо.ФизЛицо;

                                supplier.КонтактнаяИнформация.Очистить();

                                foreach (KeyValuePair<string, string> val in si.dictExInfo)
                                {
                                    if (val.Value != "")
                                    {
                                        dynamic parent = connect.Справочники.ВидыКонтактнойИнформации.НайтиПоНаименованию("Контактная информация справочника \"Контрагенты\"", true);
                                        dynamic vid = connect.Справочники.ВидыКонтактнойИнформации.НайтиПоНаименованию(val.Key, true, parent);
                                        if (vid != connect.Справочники.ВидыКонтактнойИнформации.ПустаяСсылка())
                                        {
                                            dynamic new_row = supplier.КонтактнаяИнформация.Добавить();
                                            new_row.Тип = vid.Тип;
                                            new_row.Вид = vid.Ссылка;
                                            new_row.Представление = val.Value;
                                            //if (new_row != null) Marshal.FinalReleaseComObject(new_row);
                                        }
                                        if (vid != null) Marshal.FinalReleaseComObject(vid);
                                        if (parent != null) Marshal.FinalReleaseComObject(parent);
                                    }
                                }

                                supplier.Записать();

                                findContragent = supplier.Ссылка;

                                if (supplier != null) Marshal.FinalReleaseComObject(supplier);
                                if (objects != null) Marshal.FinalReleaseComObject(objects);
                            }
                            else
                            {
                                if (objects != null) Marshal.FinalReleaseComObject(objects);
                                if (findContragent != null) Marshal.FinalReleaseComObject(findContragent);
                                if (contragent != null) Marshal.FinalReleaseComObject(contragent);
                                if (findUser != null) Marshal.FinalReleaseComObject(findUser);
                                if (findDoc != null) Marshal.FinalReleaseComObject(findDoc);
                                if (doc != null) Marshal.FinalReleaseComObject(doc);

                                backgroundWorker1.ReportProgress(1, (i + 1).ToString() + ". Документ реализации №" + listInvoiceOutDocumentsShop[i].strCode + " от " + listInvoiceOutDocumentsShop[i].date.ToShortDateString() + " пропущен, т.к. не удалось найти контрагента.");
                                continue;
                            }
                        }
                        // ------------------------

                        doc.КодДляСинхронизации = listInvoiceOutDocumentsShop[i].strCodeSync;
                        //doc.Номер = listInvoiceOutDocumentsShop[i].strCode;
                        doc.Отчет = listInvoiceOutDocumentsShop[i].bReport;
                        doc.Оптовик = listInvoiceOutDocumentsShop[i].bOptovik;
                        doc.Дата = listInvoiceOutDocumentsShop[i].date;
                        doc.Организация = ref_company.Ссылка;
                        doc.Магазин = ref_shop.Ссылка;
                        doc.Склад = ref_shop.СкладПродажи.Ссылка;
                        doc.Контрагент = findContragent;
                        if (listInvoiceOutDocumentsShop[i].bDeleted) doc.ПометкаУдаления = true;
                        doc.УчитыватьНДС = true;
                        doc.ЦенаВключаетНДС = true;
                        doc.Ответственный = findUser.Ссылка;
                        doc.Комментарий = listInvoiceOutDocumentsShop[i].strNote;
                        doc.ОтправкаПоДоговору = listInvoiceOutDocumentsShop[i].strDogovor;

                        dynamic findCard = null;
                        dynamic discountCard = null;
                        if (listInvoiceOutDocumentsShop[i].card.strName != "")
                        {
                            findCard = connect.Справочники.ИнформационныеКарты.НайтиПоНаименованию(listInvoiceOutDocumentsShop[i].card.strName, true);
                            if (findCard == connect.Справочники.ИнформационныеКарты.ПустаяСсылка())
                            {
                                // необходимо добавить контрагента и занести КодДляСинхронизации
                                discountCard = connect.Справочники.ИнформационныеКарты.СоздатьЭлемент();
                                discountCard.Наименование = listInvoiceOutDocumentsShop[i].card.strName;
                                discountCard.КодКарты = listInvoiceOutDocumentsShop[i].card.strCode;
                                discountCard.ВидКарты = connect.Перечисления.ВидыИнформационныхКарт.Магнитная;
                                discountCard.ТипКарты = connect.Перечисления.ТипыИнформационныхКарт.Дисконтная;

                                // поиск вида
                                // ------------------------
                                dynamic findType = connect.Справочники.ВидыДисконтныхКарт.НайтиПоНаименованию("Вид дисконтной карты", true);
                                dynamic type = null;
                                if (findType == connect.Справочники.ВидыДисконтныхКарт.ПустаяСсылка())
                                {
                                    // необходимо добавить тип
                                    type = connect.Справочники.ВидыДисконтныхКарт.СоздатьЭлемент();
                                    type.Наименование = "Вид дисконтной карты";
                                    type.Записать();
                                    findType = type.Ссылка;
                                }
                                // ------------------------

                                discountCard.ВидДисконтнойКарты = findType;
                                discountCard.Записать();
                                findCard = discountCard.Ссылка;

                                if (findType != null) Marshal.FinalReleaseComObject(findType);
                                if (type != null) Marshal.FinalReleaseComObject(type);
                            }
                            doc.ДисконтнаяКарта = findCard;
                        }
                        else
                        {
                            doc.ДисконтнаяКарта = connect.Справочники.ИнформационныеКарты.ПустаяСсылка();
                        }

                        doc.Товары.Очистить();

                        // добавление товаров
                        bool bError = false;

                        //double price = 0;
                        for (int j = 0; j < listInvoiceOutDocumentsShop[i].listTovars.Count(); j++)
                        {
                            dynamic findTovar = connect.Справочники.Номенклатура.НайтиПоРеквизиту("КодДляСинхронизации", listInvoiceOutDocumentsShop[i].listTovars[j].strTovarSyncCode);
                            if (findTovar == connect.Справочники.Номенклатура.ПустаяСсылка())
                            {
                                bError = true;
                                backgroundWorker1.ReportProgress(1, (i + 1).ToString() + ". Документ реализации №" + listInvoiceOutDocumentsShop[i].strCode + " от " + listInvoiceOutDocumentsShop[i].date.ToShortDateString() + " пропущен, т.к. не удалось найти номенклатуру с кодом " + listInvoiceOutDocumentsShop[i].listTovars[j].strTovarSyncCode);
                                if (findTovar != null) Marshal.FinalReleaseComObject(findTovar);
                                break;
                            }
                            dynamic findAttrib = connect.Справочники.ХарактеристикиНоменклатуры.НайтиПоРеквизиту("КодДляСинхронизации", listInvoiceOutDocumentsShop[i].listTovars[j].strAttribSyncCode);
                            if (findAttrib == connect.Справочники.ХарактеристикиНоменклатуры.ПустаяСсылка())
                            {
                                bError = true;
                                backgroundWorker1.ReportProgress(1, (i + 1).ToString() + ". Документ реализации №" + listInvoiceOutDocumentsShop[i].strCode + " от " + listInvoiceOutDocumentsShop[i].date.ToShortDateString() + " пропущен, т.к. не удалось найти характеристику номенклатуры с кодом " + listInvoiceOutDocumentsShop[i].listTovars[j].strAttribSyncCode);
                                if (findTovar != null) Marshal.FinalReleaseComObject(findTovar);
                                if (findAttrib != null) Marshal.FinalReleaseComObject(findAttrib);
                                break;
                            }

                            dynamic new_row = doc.Товары.Добавить();
                            new_row.Номенклатура = findTovar;
                            new_row.СтавкаНДС = connect.Перечисления.СтавкиНДС.БезНДС;
                            new_row.Количество = listInvoiceOutDocumentsShop[i].listTovars[j].iCount;
                            new_row.КоличествоУпаковок = listInvoiceOutDocumentsShop[i].listTovars[j].iCount;
                            new_row.Цена = listInvoiceOutDocumentsShop[i].listTovars[j].price;
                            new_row.Характеристика = findAttrib;
                            new_row.ПроцентРучнойСкидки = listInvoiceOutDocumentsShop[i].listTovars[j].manual_discount;
                            new_row.ПроцентРучнойСкидкиОпт = listInvoiceOutDocumentsShop[i].listTovars[j].manual_discount_opt;
                            //new_row.Сумма = listInvoiceOutDocumentsShop[i].listTovars[j].summa;
                            new_row.СуммаРучнойСкидки = listInvoiceOutDocumentsShop[i].listTovars[j].summa_manual_discount;

                            //price = new_row.Цена * listInvoiceOutDocumentsShop[i].listTovars[j].iCount;

                            if (new_row != null) Marshal.FinalReleaseComObject(new_row);
                            if (findTovar != null) Marshal.FinalReleaseComObject(findTovar);
                            if (findAttrib != null) Marshal.FinalReleaseComObject(findAttrib);
                        }

                        if (bError)
                        {
                            if (findContragent != null) Marshal.FinalReleaseComObject(findContragent);
                            if (contragent != null) Marshal.FinalReleaseComObject(contragent);
                            if (findCard != null) Marshal.FinalReleaseComObject(findCard);
                            if (discountCard != null) Marshal.FinalReleaseComObject(discountCard);
                            if (findUser != null) Marshal.FinalReleaseComObject(findUser);
                            if (findDoc != null) Marshal.FinalReleaseComObject(findDoc);
                            if (doc != null) Marshal.FinalReleaseComObject(doc);

                            continue;
                        }

                        //doc.СуммаДокумента = listInvoiceOutDocumentsShop[i].summa;

                        dynamic struct_ = connect.NewObject("Структура");
                        struct_.Вставить("ПрименятьКОбъекту", true);
                        struct_.Вставить("ТолькоПредварительныйРасчет", false);
                        struct_.Вставить("ВосстанавливатьУправляемыеСкидки", false);
                        struct_.Вставить("УправляемыеСкидки", null);
                        struct_.Вставить("ТолькоСообщенияПослеОформления", false);
                        struct_.Вставить("РабочееМесто", "");

                        dynamic discounts = connect.СкидкиНаценкиСерверПереопределяемый.РассчитатьПоРеализацииТоваров(doc, struct_);

                        if (doc.СуммаДокумента != listInvoiceOutDocumentsShop[i].fDocumentSum)
                        {
                            //backgroundWorker3.ReportProgress(0, "Сумма документа после применения скидок не совпадает с документом в магазине.");
                            if (MessageBox.Show("Сумма документа после применения скидок не совпадает с документом в магазине. Продолжить сохранение?", "Предупреждение", MessageBoxButtons.OKCancel) == System.Windows.Forms.DialogResult.Cancel)
                            {
                                backgroundWorker1.ReportProgress(1, (i + 1).ToString() + ". Документ реализации №" + listInvoiceOutDocumentsShop[i].strCode + " от " + listInvoiceOutDocumentsShop[i].date.ToShortDateString() + " пропущен, т.к. сумма с учетом скидок не совпадает.");
                                bError = true;
                            }
                        }

                        if (struct_ != null) Marshal.FinalReleaseComObject(struct_);
                        if (discounts != null) Marshal.FinalReleaseComObject(discounts);

                        if (!bError)
                        {
                            if (listInvoiceOutDocumentsShop[i].bProvodka)
                                doc.Записать(connect.РежимЗаписиДокумента.Проведение);
                            else
                                doc.Записать(connect.РежимЗаписиДокумента.ОтменаПроведения);
                        }

                        if (findContragent != null) Marshal.FinalReleaseComObject(findContragent);
                        if (contragent != null) Marshal.FinalReleaseComObject(contragent);
                        if (findCard != null) Marshal.FinalReleaseComObject(findCard);
                        if (discountCard != null) Marshal.FinalReleaseComObject(discountCard);
                        if (findUser != null) Marshal.FinalReleaseComObject(findUser);
                        if (findDoc != null) Marshal.FinalReleaseComObject(findDoc);
                        if (doc != null) Marshal.FinalReleaseComObject(doc);

                        backgroundWorker1.ReportProgress(0, (i + 1).ToString() + ". Документ реализации №" + listInvoiceOutDocumentsShop[i].strCode + " от " + listInvoiceOutDocumentsShop[i].date.ToShortDateString() + " успешно синхронизирован.");
                    }
                    /*catch (Exception)
                    {
                    }*/
                }

                if (listInvoiceOutDocumentsShop.Count() > 0)
                    backgroundWorker1.ReportProgress(2, "Документы реализации товаров успешно синхронизированы.");

                listInvoiceOutDocumentsShop.Clear();
            }
            catch (Exception ex)
            {
                backgroundWorker1.ReportProgress(1, "Ошибка при синхронизации документов реализации товаров: " + ex.Message);
                return false;
            }

            return true;
        }

        private bool SyncInvoiceReturn()
        {
            try
            {
                if (listInvoiceReturnDocumentsShop.Count() > 0)
                {
                    backgroundWorker1.ReportProgress(0, "");
                    backgroundWorker1.ReportProgress(0, "Синхронизация документов возврата товаров от покупателя.");
                }

                // передаем данные по приходу товаров
                for (int i = 0; i < listInvoiceReturnDocumentsShop.Count(); i++)
                {

                    // если не находим документа с таким номером и датой, просто добавляем его и проводим
                    // если находим, тогда добавляем только позиции

                    dynamic findDoc = connect.Документы.ВозвратТоваровОтПокупателя.НайтиПоРеквизиту("КодДляСинхронизации", listInvoiceReturnDocumentsShop[i].strCodeSync);
                    dynamic doc = null;
                    if (findDoc == connect.Документы.ВозвратТоваровОтПокупателя.ПустаяСсылка())
                    {
                        // если новый и не проведен, тогда не переносим
                        if (!listInvoiceReturnDocumentsShop[i].bProvodka)
                        {
                            if (findDoc != null) Marshal.FinalReleaseComObject(findDoc);
                            if (doc != null) Marshal.FinalReleaseComObject(doc);

                            backgroundWorker1.ReportProgress(0, (i + 1).ToString() + ". Документ возврата №" + listInvoiceReturnDocumentsShop[i].strCode + " от " + listInvoiceReturnDocumentsShop[i].date.ToShortDateString() + " пропущен, т.к. он не проведен.");
                            continue;
                        }
                        // необходимо добавить номенклатуру
                        doc = connect.Документы.ВозвратТоваровОтПокупателя.СоздатьДокумент();
                    }
                    else
                    {
                        // необходимо изменить документ, но он почему то не изменяется, поэтому просто удаляем этот документ и вносим новый
                        /*if (MessageBox.Show("Документ возврата №" + listInvoiceReturnDocumentsShop[i].strCode + " от " + listInvoiceReturnDocumentsShop[i].date.ToShortDateString() + " уже присутствует в основной базе. Изменить его?", "Предупреждение", MessageBoxButtons.OKCancel) == System.Windows.Forms.DialogResult.Cancel)
                        {
                            continue;
                        }*/

                        doc = findDoc.ПолучитьОбъект();
                        /*doc.Удалить();

                        if (findDoc != null) Marshal.FinalReleaseComObject(findDoc);
                        if (doc != null) Marshal.FinalReleaseComObject(doc);
                        findDoc = null;

                        doc = connect.Документы.ВозвратТоваровОтПокупателя.СоздатьДокумент();*/
                    }

                    dynamic findUser = connect.Справочники.Пользователи.НайтиПоНаименованию(listInvoiceReturnDocumentsShop[i].strUser, true);
                    if (findUser == connect.Справочники.Пользователи.ПустаяСсылка())
                    {
                        dynamic user = connect.Справочники.Пользователи.СоздатьЭлемент();
                        user.Наименование = listInvoiceReturnDocumentsShop[i].strUser;
                        user.Записать();

                        findUser = user.Ссылка;

                        if (user != null) Marshal.FinalReleaseComObject(user);
                    }

                    dynamic findContragent = connect.Справочники.Контрагенты.НайтиПоРеквизиту("КодДляСинхронизации", listInvoiceReturnDocumentsShop[i].strSupplierCodeSync);
                    dynamic contragent = null;
                    if (findContragent == connect.Справочники.Контрагенты.ПустаяСсылка())
                    {
                        // необходимо добавить контрагента

                        // получаем
                        dynamic objects = connectShop.Справочники.Контрагенты.НайтиПоРеквизиту("КодДляСинхронизации", listInvoiceReturnDocumentsShop[i].strSupplierCodeSync);

                        SupplierInfo si;
                        if (objects != connectShop.Справочники.Контрагенты.ПустаяССылка())
                        {
                            si = new SupplierInfo();
                            si.strCodeSync = objects.КодДляСинхронизации;
                            si.strName = objects.Наименование;
                            si.strFullName = objects.НаименованиеПолное;
                            si.strINN = objects.ИНН;
                            si.strNote = objects.Комментарий;

                            si.strSurname = objects.Ссылка.Фамилия;
                            si.strFirstname = objects.Ссылка.Имя;
                            si.strFathername = objects.Ссылка.Отчество;

                            si.bWholeSaler = objects.Ссылка.Оптовик;
                            si.bNoAgreeWithDelivery = objects.Ссылка.ОтказОтСМСРассылки;
                            si.bIgnor = objects.Ссылка.Игнор;

                            si.strBrandWishes = objects.Ссылка.ИнтересующиеМарки;
                            si.strSizeWishes = objects.Ссылка.ИнтересующиеРазмеры;
                            si.strCategoryWishes = objects.Ссылка.ИнтересующиеКатегории;

                            si.strWishes1 = "";
                            si.strWishes2 = "";
                            si.strWishes3 = "";

                            if (objects.Ссылка.Предпочтения1 != connectShop.Справочники.КонтрагентыПредпочтения.ПустаяСсылка())
                                si.strWishes1 = objects.Ссылка.Предпочтения1.Наименование;
                            if (objects.Ссылка.Предпочтения2 != connectShop.Справочники.КонтрагентыПредпочтения.ПустаяСсылка())
                                si.strWishes2 = objects.Ссылка.Предпочтения2.Наименование;
                            if (objects.Ссылка.Предпочтения3 != connectShop.Справочники.КонтрагентыПредпочтения.ПустаяСсылка())
                                si.strWishes3 = objects.Ссылка.Предпочтения3.Наименование;

                            si.dateChange = objects.ДатаИзменения;
                            dynamic val_ = objects.ЮрФизЛицо;
                            if (connectShop.String(val_) == connectShop.String(connectShop.Перечисления.ЮрФизЛицо.ЮрЛицо))
                                si.Type = 0;
                            else
                                si.Type = 1;

                            dynamic ex_info = objects.КонтактнаяИнформация.Выгрузить();
                            int count = ex_info.Количество();

                            for (int ii = 0; ii < count; ii++)
                            {
                                dynamic row = ex_info.Получить(ii);
                                si.dictExInfo[row.Вид.Наименование] = row.Представление;
                                Marshal.FinalReleaseComObject(row);
                            }
                            Marshal.FinalReleaseComObject(val_);

                            // добавляем
                            dynamic supplier = null;
                            supplier = connect.Справочники.Контрагенты.СоздатьЭлемент();

                            supplier.КодДляСинхронизации = si.strCodeSync;
                            supplier.Наименование = si.strName;
                            supplier.НаименованиеПолное = si.strFullName;
                            supplier.ИНН = si.strINN;
                            supplier.Комментарий = si.strNote;

                            supplier.Фамилия = si.strSurname;
                            supplier.Имя = si.strFirstname;
                            supplier.Отчество = si.strFathername;

                            supplier.Оптовик = si.bWholeSaler;
                            supplier.ОтказОтСМСРассылки = si.bNoAgreeWithDelivery;
                            supplier.Игнор = si.bIgnor;

                            supplier.ИнтересующиеМарки = si.strBrandWishes;
                            supplier.ИнтересующиеРазмеры = si.strSizeWishes;
                            supplier.ИнтересующиеКатегории = si.strCategoryWishes;

                            if (supplier.strWishes1 != "")
                            {
                                dynamic findWish = connect.Справочники.КонтрагентыПредпочтения.НайтиПоНаименованию(supplier.strWishes1, true);
                                dynamic wish = null;
                                if (findWish == connect.Справочники.КонтрагентыПредпочтения.ПустаяСсылка())
                                {
                                    // необходимо добавить тип
                                    wish = connect.Справочники.КонтрагентыПредпочтения.СоздатьЭлемент();
                                    wish.Наименование = supplier.strWishes1;
                                    wish.Записать();
                                    findWish = wish.Ссылка;
                                }
                                supplier.Предпочтения1 = findWish;

                                if (findWish != null) Marshal.FinalReleaseComObject(findWish);
                            }
                            else
                                supplier.Предпочтения1 = connect.Справочники.КонтрагентыПредпочтения.ПустаяСсылка();

                            if (supplier.strWishes2 != "")
                            {
                                dynamic findWish = connect.Справочники.КонтрагентыПредпочтения.НайтиПоНаименованию(supplier.strWishes2, true);
                                dynamic wish = null;
                                if (findWish == connect.Справочники.КонтрагентыПредпочтения.ПустаяСсылка())
                                {
                                    // необходимо добавить тип
                                    wish = connect.Справочники.КонтрагентыПредпочтения.СоздатьЭлемент();
                                    wish.Наименование = supplier.strWishes2;
                                    wish.Записать();
                                    findWish = wish.Ссылка;
                                }
                                supplier.Предпочтения2 = findWish;

                                if (findWish != null) Marshal.FinalReleaseComObject(findWish);
                            }
                            else
                                supplier.Предпочтения2 = connect.Справочники.КонтрагентыПредпочтения.ПустаяСсылка();

                            if (supplier.strWishes3 != "")
                            {
                                dynamic findWish = connect.Справочники.КонтрагентыПредпочтения.НайтиПоНаименованию(supplier.strWishes3, true);
                                dynamic wish = null;
                                if (findWish == connect.Справочники.КонтрагентыПредпочтения.ПустаяСсылка())
                                {
                                    // необходимо добавить тип
                                    wish = connect.Справочники.КонтрагентыПредпочтения.СоздатьЭлемент();
                                    wish.Наименование = supplier.strWishes3;
                                    wish.Записать();
                                    findWish = wish.Ссылка;
                                }
                                supplier.Предпочтения3 = findWish;

                                if (findWish != null) Marshal.FinalReleaseComObject(findWish);
                            }
                            else
                                supplier.Предпочтения3 = connect.Справочники.КонтрагентыПредпочтения.ПустаяСсылка();

                            if (si.Type == 0)
                                supplier.ЮрФизЛицо = connect.Перечисления.ЮрФизЛицо.ЮрЛицо;
                            else
                                supplier.ЮрФизЛицо = connect.Перечисления.ЮрФизЛицо.ФизЛицо;

                            supplier.КонтактнаяИнформация.Очистить();

                            foreach (KeyValuePair<string, string> val in si.dictExInfo)
                            {
                                if (val.Value != "")
                                {
                                    dynamic parent = connect.Справочники.ВидыКонтактнойИнформации.НайтиПоНаименованию("Контактная информация справочника \"Контрагенты\"", true);
                                    dynamic vid = connect.Справочники.ВидыКонтактнойИнформации.НайтиПоНаименованию(val.Key, true, parent);
                                    if (vid != connect.Справочники.ВидыКонтактнойИнформации.ПустаяСсылка())
                                    {
                                        dynamic new_row = supplier.КонтактнаяИнформация.Добавить();
                                        new_row.Тип = vid.Тип;
                                        new_row.Вид = vid.Ссылка;
                                        new_row.Представление = val.Value;
                                        //if (new_row != null) Marshal.FinalReleaseComObject(new_row);
                                    }
                                    if (vid != null) Marshal.FinalReleaseComObject(vid);
                                    if (parent != null) Marshal.FinalReleaseComObject(parent);
                                }
                            }

                            supplier.Записать();

                            findContragent = supplier.Ссылка;

                            if (supplier != null) Marshal.FinalReleaseComObject(supplier);
                            if (objects != null) Marshal.FinalReleaseComObject(objects);
                        }
                        else
                        {
                            if (objects != null) Marshal.FinalReleaseComObject(objects);
                            if (findContragent != null) Marshal.FinalReleaseComObject(findContragent);
                            if (contragent != null) Marshal.FinalReleaseComObject(contragent);
                            if (findUser != null) Marshal.FinalReleaseComObject(findUser);
                            if (findDoc != null) Marshal.FinalReleaseComObject(findDoc);
                            if (doc != null) Marshal.FinalReleaseComObject(doc);

                            backgroundWorker1.ReportProgress(1, (i + 1).ToString() + ". Документ возврата №" + listInvoiceReturnDocumentsShop[i].strCode + " от " + listInvoiceReturnDocumentsShop[i].date.ToShortDateString() + " пропущен, т.к. не удалось найти контрагента.");
                            continue;
                        }
                    }
                    // ------------------------

                    doc.КодДляСинхронизации = listInvoiceReturnDocumentsShop[i].strCodeSync;
                    //doc.Номер = listInvoiceReturnDocumentsShop[i].strCode;
                    doc.Отчет = listInvoiceReturnDocumentsShop[i].bReport;
                    doc.Оптовик = listInvoiceReturnDocumentsShop[i].bOptovik;
                    doc.РучнаяЦена = listInvoiceReturnDocumentsShop[i].bManualCost;
                    doc.Дата = listInvoiceReturnDocumentsShop[i].date;
                    doc.Организация = ref_company.Ссылка;
                    doc.Магазин = ref_shop.Ссылка;
                    doc.Склад = ref_shop.СкладПродажи.Ссылка;
                    doc.Контрагент = findContragent;
                    if (listInvoiceReturnDocumentsShop[i].bDeleted) doc.ПометкаУдаления = true;
                    doc.УчитыватьНДС = true;
                    doc.ЦенаВключаетНДС = true;
                    doc.Ответственный = findUser.Ссылка;
                    doc.Комментарий = listInvoiceReturnDocumentsShop[i].strNote;

                    dynamic findAnalitic = connect.Справочники.АналитикаХозяйственныхОпераций.НайтиПоНаименованию(listInvoiceReturnDocumentsShop[i].strAnaliticName, true);
                    dynamic analitic = null;
                    if (findAnalitic == connect.Справочники.АналитикаХозяйственныхОпераций.ПустаяСсылка())
                    {
                        // необходимо добавить контрагента и занести КодДляСинхронизации
                        analitic = connect.Справочники.АналитикаХозяйственныхОпераций.СоздатьЭлемент();
                        analitic.Наименование = listInvoiceReturnDocumentsShop[i].strAnaliticName;
                        analitic.ХозяйственнаяОперация = connect.Перечисления.ХозяйственныеОперации.ВозвратОтПокупателя;

                        analitic.Записать();
                        findAnalitic = analitic.Ссылка;
                    }
                    doc.АналитикаХозяйственнойОперации = findAnalitic;

                    dynamic findCard = null;
                    dynamic discountCard = null;

                    if (listInvoiceReturnDocumentsShop[i].card.strName != "")
                    {
                        findCard = connect.Справочники.ИнформационныеКарты.НайтиПоНаименованию(listInvoiceReturnDocumentsShop[i].card.strName, true);
                        if (findCard == connect.Справочники.ИнформационныеКарты.ПустаяСсылка())
                        {
                            // необходимо добавить контрагента и занести КодДляСинхронизации
                            discountCard = connect.Справочники.ИнформационныеКарты.СоздатьЭлемент();
                            discountCard.Наименование = listInvoiceReturnDocumentsShop[i].card.strName;
                            discountCard.КодКарты = listInvoiceReturnDocumentsShop[i].card.strCode;
                            discountCard.ВидКарты = connect.Перечисления.ВидыИнформационныхКарт.Магнитная;
                            discountCard.ТипКарты = connect.Перечисления.ТипыИнформационныхКарт.Дисконтная;

                            // поиск вида
                            // ------------------------
                            dynamic findType = connect.Справочники.ВидыДисконтныхКарт.НайтиПоНаименованию("Вид дисконтной карты", true);
                            dynamic type = null;
                            if (findType == connect.Справочники.ВидыДисконтныхКарт.ПустаяСсылка())
                            {
                                // необходимо добавить тип
                                type = connect.Справочники.ВидыДисконтныхКарт.СоздатьЭлемент();
                                type.Наименование = "Вид дисконтной карты";
                                type.Записать();
                                findType = type.Ссылка;
                            }
                            // ------------------------

                            discountCard.ВидДисконтнойКарты = findType;
                            discountCard.Записать();
                            findCard = discountCard.Ссылка;

                            if (findType != null) Marshal.FinalReleaseComObject(findType);
                            if (type != null) Marshal.FinalReleaseComObject(type);
                        }
                        doc.ДисконтнаяКарта = findCard;
                    }
                    else
                    {
                        doc.ДисконтнаяКарта = connect.Справочники.ИнформационныеКарты.ПустаяСсылка();
                    }

                    doc.Товары.Очистить();

                    // добавление товаров
                    bool bError = false;

                    //double price = 0;
                    for (int j = 0; j < listInvoiceReturnDocumentsShop[i].listTovars.Count(); j++)
                    {
                        dynamic findTovar = connect.Справочники.Номенклатура.НайтиПоРеквизиту("КодДляСинхронизации", listInvoiceReturnDocumentsShop[i].listTovars[j].strTovarSyncCode);
                        if (findTovar == connect.Справочники.Номенклатура.ПустаяСсылка())
                        {
                            bError = true;
                            backgroundWorker1.ReportProgress(1, (i + 1).ToString() + ". Документ возврата №" + listInvoiceReturnDocumentsShop[i].strCode + " от " + listInvoiceReturnDocumentsShop[i].date.ToShortDateString() + " пропущен, т.к. не удалось найти номенклатуру с кодом " + listInvoiceReturnDocumentsShop[i].listTovars[j].strTovarSyncCode);
                            if (findTovar != null) Marshal.FinalReleaseComObject(findTovar);
                            break;
                        }
                        dynamic findAttrib = connect.Справочники.ХарактеристикиНоменклатуры.НайтиПоРеквизиту("КодДляСинхронизации", listInvoiceReturnDocumentsShop[i].listTovars[j].strAttribSyncCode);
                        if (findAttrib == connect.Справочники.ХарактеристикиНоменклатуры.ПустаяСсылка())
                        {
                            bError = true;
                            backgroundWorker1.ReportProgress(1, (i + 1).ToString() + ". Документ возврата №" + listInvoiceReturnDocumentsShop[i].strCode + " от " + listInvoiceReturnDocumentsShop[i].date.ToShortDateString() + " пропущен, т.к. не удалось найти характеристику номенклатуры с кодом " + listInvoiceReturnDocumentsShop[i].listTovars[j].strAttribSyncCode);
                            if (findTovar != null) Marshal.FinalReleaseComObject(findTovar);
                            if (findAttrib != null) Marshal.FinalReleaseComObject(findAttrib);
                            break;
                        }
                        /*dynamic findInvoiceOut = connect.Документы.РеализацияТоваров.НайтиПоРеквизиту("КодДляСинхронизации", listInvoiceReturnDocumentsShop[i].listTovars[j].strInvoiceInCodeSync);
                        if (findInvoiceOut == connect.Документы.РеализацияТоваров.ПустаяСсылка())
                        {
                            bError = true;
                            backgroundWorker1.ReportProgress(1, (i + 1).ToString() + ". Документ возврата №" + listInvoiceReturnDocumentsShop[i].strCode + " от " + listInvoiceReturnDocumentsShop[i].date.ToShortDateString() + " пропущен, т.к. не удалось найти документ продажи с кодом " + listInvoiceOutDocumentsShop[i].listTovars[j].strInvoiceInCodeSync);
                            if (findTovar != null) Marshal.FinalReleaseComObject(findTovar);
                            if (findAttrib != null) Marshal.FinalReleaseComObject(findAttrib);
                            if (findInvoiceOut != null) Marshal.FinalReleaseComObject(findInvoiceOut);
                            break;
                        }*/

                        dynamic new_row = doc.Товары.Добавить();
                        new_row.Номенклатура = findTovar;
                        new_row.СтавкаНДС = connect.Перечисления.СтавкиНДС.БезНДС;
                        new_row.Количество = listInvoiceReturnDocumentsShop[i].listTovars[j].iCount;
                        new_row.КоличествоУпаковок = listInvoiceReturnDocumentsShop[i].listTovars[j].iCount;
                        new_row.Цена = listInvoiceReturnDocumentsShop[i].listTovars[j].price;
                        new_row.ПроцентСкидки = listInvoiceReturnDocumentsShop[i].listTovars[j].manual_discount;
                        new_row.Сумма = new_row.Цена * new_row.Количество;
                        new_row.Характеристика = findAttrib;
                        //new_row.ДокументПродажи = findInvoiceOut;

                        //price = listInvoiceReturnDocumentsShop[i].listTovars[j].price * listInvoiceReturnDocumentsShop[i].listTovars[j].iCount;

                        if (new_row != null) Marshal.FinalReleaseComObject(new_row);
                        if (findTovar != null) Marshal.FinalReleaseComObject(findTovar);
                        if (findAttrib != null) Marshal.FinalReleaseComObject(findAttrib);
                        //if (findInvoiceOut != null) Marshal.FinalReleaseComObject(findInvoiceOut);
                    }

                    if (bError)
                    {
                        if (findContragent != null) Marshal.FinalReleaseComObject(findContragent);
                        if (contragent != null) Marshal.FinalReleaseComObject(contragent);
                        if (findCard != null) Marshal.FinalReleaseComObject(findCard);
                        if (discountCard != null) Marshal.FinalReleaseComObject(discountCard);
                        if (findAnalitic != null) Marshal.FinalReleaseComObject(findAnalitic);
                        if (analitic != null) Marshal.FinalReleaseComObject(analitic);
                        if (findUser != null) Marshal.FinalReleaseComObject(findUser);
                        if (findDoc != null) Marshal.FinalReleaseComObject(findDoc);
                        if (doc != null) Marshal.FinalReleaseComObject(doc);

                        continue;
                    }

                    //doc.СуммаДокумента = listInvoiceReturnDocumentsShop[i].summa;

                    if (listInvoiceReturnDocumentsShop[i].bProvodka)
                        doc.Записать(connect.РежимЗаписиДокумента.Проведение);
                    else
                        doc.Записать(connect.РежимЗаписиДокумента.ОтменаПроведения);

                    if (findContragent != null) Marshal.FinalReleaseComObject(findContragent);
                    if (contragent != null) Marshal.FinalReleaseComObject(contragent);
                    if (findCard != null) Marshal.FinalReleaseComObject(findCard);
                    if (discountCard != null) Marshal.FinalReleaseComObject(discountCard);
                    if (findAnalitic != null) Marshal.FinalReleaseComObject(findAnalitic);
                    if (analitic != null) Marshal.FinalReleaseComObject(analitic);
                    if (findUser != null) Marshal.FinalReleaseComObject(findUser);
                    if (findDoc != null) Marshal.FinalReleaseComObject(findDoc);
                    if (doc != null) Marshal.FinalReleaseComObject(doc);

                    backgroundWorker1.ReportProgress(0, (i + 1).ToString() + ". Документ возврата №" + listInvoiceReturnDocumentsShop[i].strCode + " от " + listInvoiceReturnDocumentsShop[i].date.ToShortDateString() + " успешно синхронизирован.");
                }

                if (listInvoiceReturnDocumentsShop.Count() > 0)
                    backgroundWorker1.ReportProgress(2, "Документы возврата товаров успешно синхронизированы.");

                listInvoiceReturnDocumentsShop.Clear();
            }
            catch (Exception ex)
            {
                backgroundWorker1.ReportProgress(1, "Ошибка при синхронизации документов возврата товаров: " + ex.Message);
                return false;
            }

            return true;
        }

        private bool SyncInvoiceMoveIn()
        {
            try
            {
                if (listInvoiceMoveInDocumentsShop.Count() > 0)
                {
                    backgroundWorker1.ReportProgress(0, "");
                    backgroundWorker1.ReportProgress(0, "Синхронизация документов оприходования товаров от покупателя.");
                }

                // передаем данные по приходу товаров
                for (int i = 0; i < listInvoiceMoveInDocumentsShop.Count(); i++)
                {

                    // если не находим документа с таким номером и датой, просто добавляем его и проводим
                    // если находим, тогда добавляем только позиции

                    dynamic findDoc = connect.Документы.ОприходованиеТоваров.НайтиПоРеквизиту("КодДляСинхронизации", listInvoiceMoveInDocumentsShop[i].strCodeSync);
                    dynamic doc = null;
                    if (findDoc == connect.Документы.ОприходованиеТоваров.ПустаяСсылка())
                    {
                        // если новый и не проведен, тогда не переносим
                        if (!listInvoiceMoveInDocumentsShop[i].bProvodka)
                        {
                            if (findDoc != null) Marshal.FinalReleaseComObject(findDoc);
                            if (doc != null) Marshal.FinalReleaseComObject(doc);

                            backgroundWorker1.ReportProgress(0, (i + 1).ToString() + ". Документ оприходования №" + listInvoiceMoveInDocumentsShop[i].strCode + " от " + listInvoiceMoveInDocumentsShop[i].date.ToShortDateString() + " пропущен, т.к. он не проведен.");

                            continue;
                        }
                        doc = connect.Документы.ОприходованиеТоваров.СоздатьДокумент();
                    }
                    else
                    {
                        doc = findDoc.ПолучитьОбъект();
                    }

                    dynamic findUser = connect.Справочники.Пользователи.НайтиПоНаименованию(listInvoiceMoveInDocumentsShop[i].strUser, true);
                    if (findUser == connect.Справочники.Пользователи.ПустаяСсылка())
                    {
                        dynamic user = connect.Справочники.Пользователи.СоздатьЭлемент();
                        user.Наименование = listInvoiceMoveInDocumentsShop[i].strUser;
                        user.Записать();

                        findUser = user.Ссылка;

                        if (user != null) Marshal.FinalReleaseComObject(user);
                    }

                    // ------------------------

                    doc.КодДляСинхронизации = listInvoiceMoveInDocumentsShop[i].strCodeSync;
                    //doc.Номер = listInvoiceMoveInDocumentsShop[i].strCode;
                    doc.Дата = listInvoiceMoveInDocumentsShop[i].date;
                    doc.Организация = ref_company.Ссылка;
                    doc.Магазин = ref_shop.Ссылка;
                    doc.Склад = ref_shop.СкладПродажи.Ссылка;
                    if (listInvoiceMoveInDocumentsShop[i].bDeleted) doc.ПометкаУдаления = true;
                    doc.Ответственный = findUser.Ссылка;
                    doc.Комментарий = listInvoiceMoveInDocumentsShop[i].strNote;

                    dynamic findAnalitic = connect.Справочники.АналитикаХозяйственныхОпераций.НайтиПоНаименованию("Оприходование товаров", true);
                    dynamic analitic = null;
                    if (findAnalitic == connect.Справочники.АналитикаХозяйственныхОпераций.ПустаяСсылка())
                    {
                        // необходимо добавить контрагента и занести КодДляСинхронизации
                        analitic = connect.Справочники.АналитикаХозяйственныхОпераций.СоздатьЭлемент();
                        analitic.Наименование = "Оприходование товаров";
                        analitic.ХозяйственнаяОперация = connect.Перечисления.ХозяйственныеОперации.Оприходование;

                        analitic.Записать();
                        findAnalitic = analitic.Ссылка;
                    }
                    doc.АналитикаХозяйственнойОперации = findAnalitic;

                    doc.Товары.Очистить();

                    // добавление товаров
                    bool bError = false;

                    //double price = 0;
                    for (int j = 0; j < listInvoiceMoveInDocumentsShop[i].listTovars.Count(); j++)
                    {
                        dynamic findTovar = connect.Справочники.Номенклатура.НайтиПоРеквизиту("КодДляСинхронизации", listInvoiceMoveInDocumentsShop[i].listTovars[j].strTovarSyncCode);
                        if (findTovar == connect.Справочники.Номенклатура.ПустаяСсылка())
                        {
                            bError = true;
                            backgroundWorker1.ReportProgress(1, (i + 1).ToString() + ". Документ оприходования №" + listInvoiceMoveInDocumentsShop[i].strCode + " от " + listInvoiceMoveInDocumentsShop[i].date.ToShortDateString() + " пропущен, т.к. не удалось найти номенклатуру с кодом " + listInvoiceMoveInDocumentsShop[i].listTovars[j].strTovarSyncCode);
                            if (findTovar != null) Marshal.FinalReleaseComObject(findTovar);
                            break;
                        }
                        dynamic findAttrib = connect.Справочники.ХарактеристикиНоменклатуры.НайтиПоРеквизиту("КодДляСинхронизации", listInvoiceMoveInDocumentsShop[i].listTovars[j].strAttribSyncCode);
                        if (findAttrib == connect.Справочники.ХарактеристикиНоменклатуры.ПустаяСсылка())
                        {
                            bError = true;
                            backgroundWorker1.ReportProgress(1, (i + 1).ToString() + ". Документ оприходования №" + listInvoiceMoveInDocumentsShop[i].strCode + " от " + listInvoiceMoveInDocumentsShop[i].date.ToShortDateString() + " пропущен, т.к. не удалось найти характеристику номенклатуры с кодом " + listInvoiceMoveInDocumentsShop[i].listTovars[j].strAttribSyncCode);
                            if (findTovar != null) Marshal.FinalReleaseComObject(findTovar);
                            if (findAttrib != null) Marshal.FinalReleaseComObject(findAttrib);
                            break;
                        }

                        dynamic new_row = doc.Товары.Добавить();
                        new_row.Номенклатура = findTovar;
                        new_row.Количество = listInvoiceMoveInDocumentsShop[i].listTovars[j].iCount;
                        new_row.КоличествоУпаковок = listInvoiceMoveInDocumentsShop[i].listTovars[j].iCount;
                        new_row.Цена = listInvoiceMoveInDocumentsShop[i].listTovars[j].price;
                        new_row.Сумма = new_row.Цена * new_row.Количество;
                        new_row.Характеристика = findAttrib;

                        if (new_row != null) Marshal.FinalReleaseComObject(new_row);
                        if (findTovar != null) Marshal.FinalReleaseComObject(findTovar);
                        if (findAttrib != null) Marshal.FinalReleaseComObject(findAttrib);
                        //if (findInvoiceOut != null) Marshal.FinalReleaseComObject(findInvoiceOut);
                    }

                    if (bError)
                    {
                        if (findAnalitic != null) Marshal.FinalReleaseComObject(findAnalitic);
                        if (analitic != null) Marshal.FinalReleaseComObject(analitic);
                        if (findUser != null) Marshal.FinalReleaseComObject(findUser);
                        if (findDoc != null) Marshal.FinalReleaseComObject(findDoc);
                        if (doc != null) Marshal.FinalReleaseComObject(doc);

                        continue;
                    }

                    //doc.СуммаДокумента = listInvoiceMoveInDocumentsShop[i].summa;

                    if (listInvoiceMoveInDocumentsShop[i].bProvodka)
                        doc.Записать(connect.РежимЗаписиДокумента.Проведение);
                    else
                        doc.Записать(connect.РежимЗаписиДокумента.ОтменаПроведения);

                    if (findAnalitic != null) Marshal.FinalReleaseComObject(findAnalitic);
                    if (analitic != null) Marshal.FinalReleaseComObject(analitic);
                    if (findUser != null) Marshal.FinalReleaseComObject(findUser);
                    if (findDoc != null) Marshal.FinalReleaseComObject(findDoc);
                    if (doc != null) Marshal.FinalReleaseComObject(doc);

                    backgroundWorker1.ReportProgress(0, (i + 1).ToString() + ". Документ оприходования №" + listInvoiceMoveInDocumentsShop[i].strCode + " от " + listInvoiceMoveInDocumentsShop[i].date.ToShortDateString() + " успешно синхронизирован.");
                }

                if (listInvoiceMoveInDocumentsShop.Count() > 0)
                    backgroundWorker1.ReportProgress(2, "Документы оприходования товаров успешно синхронизированы.");

                listInvoiceMoveInDocumentsShop.Clear();
            }
            catch (Exception ex)
            {
                backgroundWorker1.ReportProgress(1, "Ошибка при синхронизации документов оприходования: " + ex.Message);
                return false;
            }

            return true;
        }

        private bool SyncInvoiceMoveOut()
        {
            try
            {
                if (listInvoiceMoveOutDocumentsShop.Count() > 0)
                {
                    backgroundWorker1.ReportProgress(0, "");
                    backgroundWorker1.ReportProgress(0, "Синхронизация документов списания товаров от покупателя.");
                }

                // передаем данные по приходу товаров
                for (int i = 0; i < listInvoiceMoveOutDocumentsShop.Count(); i++)
                {

                    // если не находим документа с таким номером и датой, просто добавляем его и проводим
                    // если находим, тогда добавляем только позиции

                    dynamic findDoc = connect.Документы.СписаниеТоваров.НайтиПоРеквизиту("КодДляСинхронизации", listInvoiceMoveOutDocumentsShop[i].strCodeSync);
                    dynamic doc = null;
                    if (findDoc == connect.Документы.СписаниеТоваров.ПустаяСсылка())
                    {
                        // если новый и не проведен, тогда не переносим
                        if (!listInvoiceMoveOutDocumentsShop[i].bProvodka)
                        {
                            if (findDoc != null) Marshal.FinalReleaseComObject(findDoc);
                            if (doc != null) Marshal.FinalReleaseComObject(doc);

                            backgroundWorker1.ReportProgress(0, (i + 1).ToString() + ". Документ списания №" + listInvoiceMoveOutDocumentsShop[i].strCode + " от " + listInvoiceMoveOutDocumentsShop[i].date.ToShortDateString() + " пропущен, т.к. он не проведен.");

                            continue;
                        }
                        doc = connect.Документы.СписаниеТоваров.СоздатьДокумент();
                    }
                    else
                    {
                        doc = findDoc.ПолучитьОбъект();
                    }

                    dynamic findUser = connect.Справочники.Пользователи.НайтиПоНаименованию(listInvoiceMoveOutDocumentsShop[i].strUser, true);
                    if (findUser == connect.Справочники.Пользователи.ПустаяСсылка())
                    {
                        dynamic user = connect.Справочники.Пользователи.СоздатьЭлемент();
                        user.Наименование = listInvoiceMoveOutDocumentsShop[i].strUser;
                        user.Записать();

                        findUser = user.Ссылка;

                        if (user != null) Marshal.FinalReleaseComObject(user);
                    }

                    // ------------------------

                    doc.КодДляСинхронизации = listInvoiceMoveOutDocumentsShop[i].strCodeSync;
                    //doc.Номер = listInvoiceMoveOutDocumentsShop[i].strCode;
                    doc.Дата = listInvoiceMoveOutDocumentsShop[i].date;
                    doc.Организация = ref_company.Ссылка;
                    doc.Магазин = ref_shop.Ссылка;
                    doc.Склад = ref_shop.СкладПродажи.Ссылка;
                    if (listInvoiceMoveOutDocumentsShop[i].bDeleted) doc.ПометкаУдаления = true;
                    doc.Ответственный = findUser.Ссылка;
                    doc.Комментарий = listInvoiceMoveOutDocumentsShop[i].strNote;

                    dynamic findAnalitic = connect.Справочники.АналитикаХозяйственныхОпераций.НайтиПоНаименованию("Списание на затраты (подарки)", true);
                    dynamic analitic = null;
                    if (findAnalitic == connect.Справочники.АналитикаХозяйственныхОпераций.ПустаяСсылка())
                    {
                        // необходимо добавить контрагента и занести КодДляСинхронизации
                        analitic = connect.Справочники.АналитикаХозяйственныхОпераций.СоздатьЭлемент();
                        analitic.Наименование = "Списание на затраты (подарки)";
                        analitic.ХозяйственнаяОперация = connect.Перечисления.ХозяйственныеОперации.СписаниеНаЗатраты;

                        analitic.Записать();
                        findAnalitic = analitic.Ссылка;
                    }
                    doc.АналитикаХозяйственнойОперации = findAnalitic;

                    doc.Товары.Очистить();

                    // добавление товаров
                    bool bError = false;

                    //double price = 0;
                    for (int j = 0; j < listInvoiceMoveOutDocumentsShop[i].listTovars.Count(); j++)
                    {
                        dynamic findTovar = connect.Справочники.Номенклатура.НайтиПоРеквизиту("КодДляСинхронизации", listInvoiceMoveOutDocumentsShop[i].listTovars[j].strTovarSyncCode);
                        if (findTovar == connect.Справочники.Номенклатура.ПустаяСсылка())
                        {
                            bError = true;
                            backgroundWorker1.ReportProgress(1, (i + 1).ToString() + ". Документ списания №" + listInvoiceMoveOutDocumentsShop[i].strCode + " от " + listInvoiceMoveOutDocumentsShop[i].date.ToShortDateString() + " пропущен, т.к. не удалось найти номенклатуру с кодом " + listInvoiceMoveOutDocumentsShop[i].listTovars[j].strTovarSyncCode);
                            if (findTovar != null) Marshal.FinalReleaseComObject(findTovar);
                            break;
                        }
                        dynamic findAttrib = connect.Справочники.ХарактеристикиНоменклатуры.НайтиПоРеквизиту("КодДляСинхронизации", listInvoiceMoveOutDocumentsShop[i].listTovars[j].strAttribSyncCode);
                        if (findAttrib == connect.Справочники.ХарактеристикиНоменклатуры.ПустаяСсылка())
                        {
                            bError = true;
                            backgroundWorker1.ReportProgress(1, (i + 1).ToString() + ". Документ списания №" + listInvoiceMoveOutDocumentsShop[i].strCode + " от " + listInvoiceMoveOutDocumentsShop[i].date.ToShortDateString() + " пропущен, т.к. не удалось найти характеристику номенклатуры с кодом " + listInvoiceMoveOutDocumentsShop[i].listTovars[j].strAttribSyncCode);
                            if (findTovar != null) Marshal.FinalReleaseComObject(findTovar);
                            if (findAttrib != null) Marshal.FinalReleaseComObject(findAttrib);
                            break;
                        }

                        dynamic new_row = doc.Товары.Добавить();
                        new_row.Номенклатура = findTovar;
                        new_row.Количество = listInvoiceMoveOutDocumentsShop[i].listTovars[j].iCount;
                        new_row.КоличествоУпаковок = listInvoiceMoveOutDocumentsShop[i].listTovars[j].iCount;
                        new_row.Цена = listInvoiceMoveOutDocumentsShop[i].listTovars[j].price;
                        new_row.Сумма = new_row.Цена * new_row.Количество;
                        new_row.Характеристика = findAttrib;

                        if (new_row != null) Marshal.FinalReleaseComObject(new_row);
                        if (findTovar != null) Marshal.FinalReleaseComObject(findTovar);
                        if (findAttrib != null) Marshal.FinalReleaseComObject(findAttrib);
                        //if (findInvoiceOut != null) Marshal.FinalReleaseComObject(findInvoiceOut);
                    }

                    if (bError)
                    {
                        if (findAnalitic != null) Marshal.FinalReleaseComObject(findAnalitic);
                        if (analitic != null) Marshal.FinalReleaseComObject(analitic);
                        if (findUser != null) Marshal.FinalReleaseComObject(findUser);
                        if (findDoc != null) Marshal.FinalReleaseComObject(findDoc);
                        if (doc != null) Marshal.FinalReleaseComObject(doc);

                        continue;
                    }

                    //doc.СуммаДокумента = listInvoiceMoveOutDocumentsShop[i].summa;

                    if (listInvoiceMoveOutDocumentsShop[i].bProvodka)
                        doc.Записать(connect.РежимЗаписиДокумента.Проведение);
                    else
                        doc.Записать(connect.РежимЗаписиДокумента.ОтменаПроведения);

                    if (findAnalitic != null) Marshal.FinalReleaseComObject(findAnalitic);
                    if (analitic != null) Marshal.FinalReleaseComObject(analitic);
                    if (findUser != null) Marshal.FinalReleaseComObject(findUser);
                    if (findDoc != null) Marshal.FinalReleaseComObject(findDoc);
                    if (doc != null) Marshal.FinalReleaseComObject(doc);

                    backgroundWorker1.ReportProgress(0, (i + 1).ToString() + ". Документ списания №" + listInvoiceMoveOutDocumentsShop[i].strCode + " от " + listInvoiceMoveOutDocumentsShop[i].date.ToShortDateString() + " успешно синхронизирован.");
                }

                if (listInvoiceMoveOutDocumentsShop.Count() > 0)
                    backgroundWorker1.ReportProgress(2, "Документы списания товаров успешно синхронизированы.");

                listInvoiceMoveOutDocumentsShop.Clear();
            }
            catch (Exception ex)
            {
                backgroundWorker1.ReportProgress(1, "Ошибка при синхронизации документов списания: " + ex.Message);
                return false;
            }

            return true;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            gbMain.Enabled = false;
            gbShop.Enabled = false;
            gbSync.Enabled = false;
            btnRefreshDataSite.Enabled = false;
            btnYML.Enabled = false;

            listView1.Items.Clear();

            m_strSelectedShop = cbShop.Items[cbShop.SelectedIndex].ToString();

            if (ref_shop != null) Marshal.FinalReleaseComObject(ref_shop);
            if (ref_company != null) Marshal.FinalReleaseComObject(ref_company);

            ref_shop = connect.Справочники.Магазины.НайтиПоНаименованию(m_strSelectedShop, true);
            if (ref_shop == connect.Справочники.Магазины.ПустаяСсылка())
                MessageBox.Show("Необходимо выбрать магазин.");
            // проверяем, чтобы была хотя бы одна организация
            dynamic query = connect.NewObject("Запрос");
            query.Текст = "ВЫБРАТЬ первые 1 Ссылка ИЗ Справочник.Организации ГДЕ ПометкаУдаления = ложь";
            dynamic objects = query.Выполнить().Выбрать();
            while (objects.Следующий())
            {
                ref_company = objects.Ссылка;
                break;
            }
            if (ref_company == null)
                MessageBox.Show("Необходимо, чтобы в основной базе существовала хотя бы одна запись в справочнике организаций.");

            if (query != null) Marshal.FinalReleaseComObject(query);
            if (objects != null) Marshal.FinalReleaseComObject(objects);

            backgroundWorker1.RunWorkerAsync();
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            SyncAll();
        }

        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            if (e.ProgressPercentage == -1)
            {
                ListViewItem vi = listView1.Items[listView1.Items.Count - 1];
                vi.Text = (string)e.UserState;
                listView1.EnsureVisible(listView1.Items.Count - 1);
            }
            else
            {
                ListViewItem vi = listView1.Items.Add((string)e.UserState);
                if ((string)e.UserState != "") vi.ImageIndex = e.ProgressPercentage;
                listView1.EnsureVisible(listView1.Items.Count - 1);
            }
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (listView1.Items.Count > 0)
                listView1.EnsureVisible(listView1.Items.Count - 1);

            gbMain.Enabled = true;
            gbShop.Enabled = true;
            gbSync.Enabled = true;
            btnRefreshDataSite.Enabled = true;
            btnYML.Enabled = true;

            tbTovarChange.Text = "";
            tbTovarChange2.Text = "";

            tbPriceChange.Text = "";
            tbPriceChange2.Text = "";
            tbDiscountChange.Text = "";
            //tbInvoiceInChange.Text = "";
            tbInvoiceInChange2.Text = "";
            tbInvoiceOutChange2.Text = "";
            tbInvoiceReturnChange2.Text = "";
            tbInvoiceMoveInChange2.Text = "";
            tbInvoiceMoveOutChange2.Text = "";

            tbContragentChange.Text = "";
            tbContragentChange2.Text = "";

            if (listTovars.Count() > 0) tbTovarChange.Text = listTovars.Count().ToString();
            if (listTovarsShop.Count() > 0) tbTovarChange2.Text = listTovarsShop.Count().ToString();

            if (listBrands.Count() > 0)
            {
                if (tbTovarChange.Text == "") tbTovarChange.Text = "0 / " + listBrands.Count().ToString();
                else tbTovarChange.Text += " / " + listBrands.Count().ToString();
            }
            if (listBrandsShop.Count() > 0)
            {
                if (tbTovarChange2.Text == "") tbTovarChange2.Text = "0 / " + listBrandsShop.Count().ToString();
                else tbTovarChange2.Text += " / " + listBrandsShop.Count().ToString();
            }

            if (listPriceDocuments.Count() > 0) tbPriceChange.Text = listPriceDocuments.Count().ToString();
            if (listPriceDocumentsShop.Count() > 0) tbPriceChange2.Text = listPriceDocumentsShop.Count().ToString();
            if (listSuppliers.Count() > 0) tbContragentChange.Text = listSuppliers.Count().ToString();
            if (listSuppliersShop.Count() > 0) tbContragentChange2.Text = listSuppliersShop.Count().ToString();

            // сравниваем акции
            if (ActionDocument == null && ActionShopDocument != null ||
                ActionDocument != null && ActionShopDocument == null ||
                ActionDocument != null && ActionShopDocument != null &&
                (ActionDocument.bDeleted != ActionShopDocument.bDeleted ||
                ActionDocument.bProvodka != ActionShopDocument.bProvodka ||
                ActionDocument.dateBegin != ActionShopDocument.dateBegin ||
                ActionDocument.dateEnd != ActionShopDocument.dateEnd ||
                ActionDocument.strName != ActionShopDocument.strName ||
                ActionDocument.listDiscounts.Count != ActionShopDocument.listDiscounts.Count
                || !EqualActionDiscounts()))
            {
                // нужно синхронизировать
                tbDiscountChange.Text = "есть изменения";
            }

            //if (listInvoiceInDocuments.Count > 0) tbInvoiceInChange.Text = listInvoiceInDocuments.Count().ToString();
            if (listInvoiceInDocumentsShop.Count > 0) tbInvoiceInChange2.Text = listInvoiceInDocumentsShop.Count().ToString();
            if (listInvoiceOutDocumentsShop.Count > 0) tbInvoiceOutChange2.Text = listInvoiceOutDocumentsShop.Count().ToString();
            if (listInvoiceReturnDocumentsShop.Count > 0) tbInvoiceReturnChange2.Text = listInvoiceReturnDocumentsShop.Count().ToString();
            if (listInvoiceMoveInDocumentsShop.Count > 0) tbInvoiceMoveInChange2.Text = listInvoiceMoveInDocumentsShop.Count().ToString();
            if (listInvoiceMoveOutDocumentsShop.Count > 0) tbInvoiceMoveOutChange2.Text = listInvoiceMoveOutDocumentsShop.Count().ToString();

            if (listInvoiceInDocumentsShop.Count == 0 &&
                listInvoiceOutDocumentsShop.Count == 0 &&
                listInvoiceReturnDocumentsShop.Count == 0 &&
                listInvoiceMoveInDocumentsShop.Count == 0 &&
                listInvoiceMoveOutDocumentsShop.Count == 0)
            {
                btnFixRemains.Enabled = true;
                cbDeleteAutoInvoice.Enabled = true;
                cbVerifyMainBase.Enabled = true;
            }

        }

        private void btnFixRemains_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Зафиксировать остатки на складе?", "Предупреждение", MessageBoxButtons.OKCancel) != System.Windows.Forms.DialogResult.OK)
                return;

            gbMain.Enabled = false;
            gbShop.Enabled = false;
            gbSync.Enabled = false;
            btnRefreshDataSite.Enabled = false;
            btnYML.Enabled = false;

            listView1.Items.Clear();

            m_strSelectedShop = cbShop.Items[cbShop.SelectedIndex].ToString();

            if (cbDeleteAutoInvoice.Checked) m_bDeleteAutoInvoice = true;
            else m_bDeleteAutoInvoice = false;

            if (ref_shop != null) Marshal.FinalReleaseComObject(ref_shop);
            //if (ref_company != null) Marshal.FinalReleaseComObject(ref_company);

            ref_shop = connect.Справочники.Магазины.НайтиПоНаименованию(m_strSelectedShop, true);
            if (ref_shop == connect.Справочники.Магазины.ПустаяСсылка())
                MessageBox.Show("Необходимо выбрать магазин.");
            // проверяем, чтобы была хотя бы одна организация
            /*dynamic query = connect.NewObject("Запрос");
            query.Текст = "ВЫБРАТЬ первые 1 Ссылка ИЗ Справочник.Организации ГДЕ ПометкаУдаления = ложь";
            dynamic objects = query.Выполнить().Выбрать();
            while (objects.Следующий())
            {
                ref_company = objects.Ссылка;
                break;
            }
            if (ref_company == null)
                MessageBox.Show("Необходимо, чтобы в основной базе существовала хотя бы одна запись в справочнике организаций.");*/

            //if (query != null) Marshal.FinalReleaseComObject(query);
            //if (objects != null) Marshal.FinalReleaseComObject(objects);

            if (cbVerifyMainBase.Checked) m_bVerifyMainBase = true;
            else m_bVerifyMainBase = false;

            if (cbNoDeleteDocuments.Checked) m_bNoDeleteDocuments = true;
            else m_bNoDeleteDocuments = false;

            backgroundWorker2.RunWorkerAsync();

        }

        private void backgroundWorker2_DoWork(object sender, DoWorkEventArgs e)
        {
            SyncRemains();
        }

        private void backgroundWorker2_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            if (e.ProgressPercentage == -1)
            {
                ListViewItem vi = listView1.Items[listView1.Items.Count - 1];
                vi.Text = (string)e.UserState;
                listView1.EnsureVisible(listView1.Items.Count - 1);
            }
            else
            {
                ListViewItem vi = listView1.Items.Add((string)e.UserState);
                if ((string)e.UserState != "") vi.ImageIndex = e.ProgressPercentage;
                listView1.EnsureVisible(listView1.Items.Count - 1);
            }
        }

        private void backgroundWorker2_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (listView1.Items.Count > 0)
                listView1.EnsureVisible(listView1.Items.Count - 1);

            gbMain.Enabled = true;
            gbShop.Enabled = true;
            gbSync.Enabled = true;
            btnRefreshDataSite.Enabled = true;
            btnYML.Enabled = true;
        }

        private void cbDeleteAutoInvoice_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                if (!m_bSaveInProgress)
                    connectShop.Константы.УдалятьПредыдущиеФиктивныеДокументы.Установить(cbDeleteAutoInvoice.Checked);
            }
            catch (Exception)
            {
                m_bSaveInProgress = true;
                cbDeleteAutoInvoice.Checked = !cbDeleteAutoInvoice.Checked;
                m_bSaveInProgress = false;

                ListViewItem vi = listView1.Items.Add("Не удалось сохранить значение в базе магазина.");
                vi.ImageIndex = 1;
                listView1.EnsureVisible(listView1.Items.Count - 1);
            }

        }

        private void backgroundWorker3_DoWork(object sender, DoWorkEventArgs e)
        {
            GetData();
        }

        private void backgroundWorker3_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            if (e.ProgressPercentage == -1)
            {
                ListViewItem vi = listView1.Items[listView1.Items.Count - 1];
                vi.Text = (string)e.UserState;
                listView1.EnsureVisible(listView1.Items.Count - 1);
            }
            else
            {
                ListViewItem vi = listView1.Items.Add((string)e.UserState);
                if ((string)e.UserState != "") vi.ImageIndex = e.ProgressPercentage;
                listView1.EnsureVisible(listView1.Items.Count - 1);
            }
        }

        private void backgroundWorker3_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (listView1.Items.Count > 0)
                listView1.EnsureVisible(listView1.Items.Count - 1);

            tbTovarChange.Text = "";
            if (listTovars.Count() > 0) tbTovarChange.Text = listTovars.Count().ToString();

            tbTovarChange2.Text = "";
            if (listTovarsShop.Count() > 0) tbTovarChange2.Text = listTovarsShop.Count().ToString();

            if (listBrands.Count() > 0)
            {
                if (tbTovarChange.Text == "") tbTovarChange.Text = "0 / " + listBrands.Count().ToString();
                else tbTovarChange.Text += " / " + listBrands.Count().ToString();
            }
            if (listBrandsShop.Count() > 0)
            {
                if (tbTovarChange2.Text == "") tbTovarChange2.Text = "0 / " + listBrandsShop.Count().ToString();
                else tbTovarChange2.Text += " / " + listBrandsShop.Count().ToString();
            }

            tbPriceChange.Text = "";
            if (listPriceDocuments.Count() > 0) tbPriceChange.Text = listPriceDocuments.Count().ToString();

            tbPriceChange2.Text = "";
            if (listPriceDocumentsShop.Count() > 0) tbPriceChange2.Text = listPriceDocumentsShop.Count().ToString();

            tbDiscountChange.Text = "";
            // сравниваем акции
            if (ActionDocument == null && ActionShopDocument != null ||
                ActionDocument != null && ActionShopDocument == null ||
                ActionDocument != null && ActionShopDocument != null &&
                (ActionDocument.bDeleted != ActionShopDocument.bDeleted ||
                ActionDocument.bProvodka != ActionShopDocument.bProvodka ||
                ActionDocument.dateBegin != ActionShopDocument.dateBegin ||
                ActionDocument.dateEnd != ActionShopDocument.dateEnd ||
                ActionDocument.strName != ActionShopDocument.strName ||
                ActionDocument.listDiscounts.Count != ActionShopDocument.listDiscounts.Count
                || !EqualActionDiscounts()))
            {
                // нужно синхронизировать
                tbDiscountChange.Text = "есть изменения";
            }

            tbContragentChange.Text = "";
            if (listSuppliers.Count() > 0) tbContragentChange.Text = listSuppliers.Count().ToString();

            tbContragentChange2.Text = "";
            if (listSuppliersShop.Count() > 0) tbContragentChange2.Text = listSuppliersShop.Count().ToString();

            tbInvoiceInChange2.Text = "";
            if (listInvoiceInDocumentsShop.Count() > 0) tbInvoiceInChange2.Text = listInvoiceInDocumentsShop.Count().ToString();

            tbInvoiceOutChange2.Text = "";
            if (listInvoiceOutDocumentsShop.Count > 0) tbInvoiceOutChange2.Text = listInvoiceOutDocumentsShop.Count().ToString();

            tbInvoiceReturnChange2.Text = "";
            if (listInvoiceReturnDocumentsShop.Count > 0) tbInvoiceReturnChange2.Text = listInvoiceReturnDocumentsShop.Count().ToString();

            tbInvoiceMoveInChange2.Text = "";
            if (listInvoiceMoveInDocumentsShop.Count > 0) tbInvoiceMoveInChange2.Text = listInvoiceMoveInDocumentsShop.Count().ToString();

            tbInvoiceMoveOutChange2.Text = "";
            if (listInvoiceMoveOutDocumentsShop.Count > 0) tbInvoiceMoveOutChange2.Text = listInvoiceMoveOutDocumentsShop.Count().ToString();

            if (listInvoiceInDocumentsShop.Count == 0 &&
                listInvoiceOutDocumentsShop.Count == 0 &&
                listInvoiceReturnDocumentsShop.Count == 0 &&
                listInvoiceMoveInDocumentsShop.Count == 0 &&
                listInvoiceMoveOutDocumentsShop.Count == 0)
            {
                btnFixRemains.Enabled = true;
                cbDeleteAutoInvoice.Enabled = true;
                cbVerifyMainBase.Enabled = true;
            }

            bool bVal1 = connectShop.Константы.УдалятьПредыдущиеФиктивныеДокументы.Получить();
            m_bSaveInProgress = true;
            cbDeleteAutoInvoice.Checked = bVal1;
            m_bSaveInProgress = false;

            button1.Enabled = true;
            gbSync.Enabled = true;
            btnRefreshDataSite.Enabled = true;
            btnYML.Enabled = true;
        }

        void GetData()
        {
            try
            {
                // -------
                // товары
                // -------
                DateTime dateSyncShop = connectShop.Константы.ДатаСинхронизации.Получить();

                {
                    Cursor.Current = Cursors.WaitCursor;

                    for (int i = 0; i < listTovars.Count(); i++)
                    {
                        Marshal.FinalReleaseComObject(listTovars[i].ref_);
                    }

                    listTovars.Clear();

                    for (int i = 0; i < listTovarsShop.Count(); i++)
                    {
                        Marshal.FinalReleaseComObject(listTovarsShop[i].ref_);
                    }

                    listTovarsShop.Clear();

                    backgroundWorker3.ReportProgress(0, "Обновление даты изменения номенклатур в базе магазина.");
                    dynamic mod2 = connectShop.ExchangeData;
                    int res = mod2.ОбновитьДатуИзмененияНоменклатур(dateSyncShop);
                    Marshal.FinalReleaseComObject(mod2);
                    backgroundWorker3.ReportProgress(-1, "Обновление даты изменения номенклатур в базе магазина (ok).");

                    backgroundWorker3.ReportProgress(0, "Обновление даты изменения номенклатур в основной базе.");
                    dynamic mod = connect.ExchangeData;
                    res = mod.ОбновитьДатуИзмененияНоменклатур(DateSync);
                    Marshal.FinalReleaseComObject(mod);
                    backgroundWorker3.ReportProgress(-1, "Обновление даты изменения номенклатур в основной базе (ok).");

                    {
                        backgroundWorker3.ReportProgress(0, "Получение данных по номенклатурам из основной базы.");

                        // номенклатуры из основной базы
                        dynamic query = connect.NewObject("Запрос");
                        query.Текст = "ВЫБРАТЬ Код, КодДляСинхронизации, Ссылка, ДатаИзменения, ДатаИзмененияИзображений, ДатаИзмененияАттрибутов ИЗ Справочник.Номенклатура ГДЕ ЭтоГруппа = ложь "
                        + " И (ДатаИзменения > &Data1 ИЛИ ДатаИзмененияИзображений > &Data2 ИЛИ ДатаИзмененияАттрибутов > &Data3)";

                        query.УстановитьПараметр("Data1", DateSync);
                        query.УстановитьПараметр("Data2", DateSync);
                        query.УстановитьПараметр("Data3", DateSync);

                        dynamic objects2 = query.Выполнить();
                        dynamic objects = objects2.Выбрать();

                        TovarInfo ti;
                        while (objects.Следующий())
                        {
                            ti = new TovarInfo();
                            ti.ref_ = objects.Ссылка;
                            ti.strCodeSync = objects.КодДляСинхронизации;
                            ti.dateChange = objects.ДатаИзменения;
                            ti.dateChangeImages = objects.ДатаИзмененияИзображений;
                            ti.dateChangeAttribs = objects.ДатаИзмененияАттрибутов;
                            ti.strArticul = objects.Ссылка.Артикул;
                            ti.strFullName = objects.Ссылка.НаименованиеПолное;
                            ti.strName = objects.Ссылка.Наименование;
                            ti.bDeleted = objects.Ссылка.ПометкаУдаления;

                            ti.strNote = objects.Ссылка.Описание;
                            ti.strNameForSite = objects.Ссылка.НаименованиеДляСайта;
                            ti.strNoteForSite = objects.Ссылка.ПримечаниеДляСайта;

                            ti.fDiscount = objects.Ссылка.СкидкаРозн;
                            ti.bBigSize = objects.Ссылка.ДляПолных;
                            ti.iSort = objects.Ссылка.Сортировка;
                            ti.bNew = objects.Ссылка.Новинка;
                            ti.fPriceInEuro = objects.Ссылка.ЦенаЕвро;
                            ti.bInvisible = objects.Ссылка.Невидимость;

                            ti.group.strCodeSync = objects.Ссылка.Родитель.КодДляСинхронизации;
                            ti.group.strName = objects.Ссылка.Родитель.Наименование;
                            ti.group.bInvisible = objects.Ссылка.Родитель.Невидимость;
                            ti.group.bOnlyForRegistered = objects.Ссылка.Родитель.ЦенаТолькоДляЗарегистрированныхПользователей;
                            ti.group.bNewCollection = objects.Ссылка.Родитель.НоваяКоллекция;
                            ti.group.fDiscount = objects.Ссылка.Родитель.СкидкаРозн;
                            ti.group.fDiscountWholesale = objects.Ссылка.Родитель.СкидкаОпт;

                            ti.group.brand.strCodeSync = objects.Ссылка.Родитель.ор_ТорговаяМарка.КодДляСинхронизации;
                            ti.group.brand.strName = objects.Ссылка.Родитель.ор_ТорговаяМарка.Наименование;
                            ti.group.brand.iSortPos = objects.Ссылка.Родитель.ор_ТорговаяМарка.Сортировка;

                            ti.group.collection.strCodeSync = objects.Ссылка.Родитель.ор_Сезон.КодДляСинхронизации;
                            ti.group.collection.strName = objects.Ссылка.Родитель.ор_Сезон.Наименование;
                            ti.group.collection.bInvisible = objects.Ссылка.Родитель.ор_Сезон.Невидимость;
                            ti.group.collection.iSortPos = objects.Ссылка.Родитель.ор_Сезон.Сортировка;

                            ti.group.collectionForSite.strCodeSync = objects.Ссылка.Родитель.op_СезонДляСайта.КодДляСинхронизации;
                            ti.group.collectionForSite.strName = objects.Ссылка.Родитель.op_СезонДляСайта.Наименование;
                            ti.group.collectionForSite.bInvisible = objects.Ссылка.Родитель.op_СезонДляСайта.Невидимость;
                            ti.group.collectionForSite.iSortPos = objects.Ссылка.Родитель.op_СезонДляСайта.Сортировка;

                            ti.type.strCodeSync = objects.Ссылка.ВидНоменклатуры.КодДляСинхронизации;
                            ti.type.strName = objects.Ссылка.ВидНоменклатуры.Наименование;
                            ti.type.bInvisible = objects.Ссылка.ВидНоменклатуры.Невидимость;

                            listTovars.Add(ti);

                            backgroundWorker3.ReportProgress(-1, "Получение данных по номенклатурам из основной базы (" + listTovars.Count.ToString() + ").");
                        }

                        Marshal.FinalReleaseComObject(objects2);
                        Marshal.FinalReleaseComObject(objects);
                        Marshal.FinalReleaseComObject(query);
                    }

                    // номенклатуры из базы магазина
                    {
                        backgroundWorker3.ReportProgress(0, "Получение данных по номенклатурам из базы магазина.");

                        dynamic queryShop = connectShop.NewObject("Запрос");
                        queryShop.Текст = "ВЫБРАТЬ Код, КодДляСинхронизации, Ссылка, ДатаИзменения, ДатаИзмененияИзображений, ДатаИзмененияАттрибутов ИЗ Справочник.Номенклатура ГДЕ ЭтоГруппа = ложь "
                        + " И (ДатаИзменения > &Data1 ИЛИ ДатаИзмененияИзображений > &Data2 ИЛИ ДатаИзмененияАттрибутов > &Data3)";

                        //DateTime dateSyncShop = connectShop.Константы.ДатаСинхронизации.Получить();

                        queryShop.УстановитьПараметр("Data1", dateSyncShop);
                        queryShop.УстановитьПараметр("Data2", dateSyncShop);
                        queryShop.УстановитьПараметр("Data3", dateSyncShop);

                        dynamic objectsShop2 = queryShop.Выполнить();
                        dynamic objectsShop = objectsShop2.Выбрать();

                        TovarInfo ti;
                        while (objectsShop.Следующий())
                        {
                            ti = new TovarInfo();
                            ti.ref_ = objectsShop.Ссылка;
                            ti.strCodeSync = objectsShop.КодДляСинхронизации;
                            ti.dateChange = objectsShop.ДатаИзменения;
                            ti.dateChangeImages = objectsShop.ДатаИзмененияИзображений;
                            ti.dateChangeAttribs = objectsShop.ДатаИзмененияАттрибутов;
                            ti.strArticul = objectsShop.Ссылка.Артикул;
                            ti.strFullName = objectsShop.Ссылка.НаименованиеПолное;
                            ti.strName = objectsShop.Ссылка.Наименование;
                            ti.bDeleted = objectsShop.Ссылка.ПометкаУдаления;

                            ti.strNote = objectsShop.Ссылка.Описание;
                            ti.strNameForSite = objectsShop.Ссылка.НаименованиеДляСайта;
                            ti.strNoteForSite = objectsShop.Ссылка.ПримечаниеДляСайта;

                            ti.fDiscount = objectsShop.Ссылка.СкидкаРозн;
                            ti.bBigSize = objectsShop.Ссылка.ДляПолных;
                            ti.iSort = objectsShop.Ссылка.Сортировка;
                            ti.bNew = objectsShop.Ссылка.Новинка;
                            ti.fPriceInEuro = objectsShop.Ссылка.ЦенаЕвро;
                            ti.bInvisible = objectsShop.Ссылка.Невидимость;

                            ti.group.strCodeSync = objectsShop.Ссылка.Родитель.КодДляСинхронизации;
                            ti.group.strName = objectsShop.Ссылка.Родитель.Наименование;
                            ti.group.bInvisible = objectsShop.Ссылка.Родитель.Невидимость;
                            ti.group.bOnlyForRegistered = objectsShop.Ссылка.Родитель.ЦенаТолькоДляЗарегистрированныхПользователей;
                            ti.group.bNewCollection = objectsShop.Ссылка.Родитель.НоваяКоллекция;
                            ti.group.fDiscount = objectsShop.Ссылка.Родитель.СкидкаРозн;
                            ti.group.fDiscountWholesale = objectsShop.Ссылка.Родитель.СкидкаОпт;

                            ti.group.brand.strCodeSync = objectsShop.Ссылка.Родитель.ор_ТорговаяМарка.КодДляСинхронизации;
                            ti.group.brand.strName = objectsShop.Ссылка.Родитель.ор_ТорговаяМарка.Наименование;
                            ti.group.brand.iSortPos = objectsShop.Ссылка.Родитель.ор_ТорговаяМарка.Сортировка;

                            ti.group.collection.strCodeSync = objectsShop.Ссылка.Родитель.ор_Сезон.КодДляСинхронизации;
                            ti.group.collection.strName = objectsShop.Ссылка.Родитель.ор_Сезон.Наименование;
                            ti.group.collection.bInvisible = objectsShop.Ссылка.Родитель.ор_Сезон.Невидимость;
                            ti.group.collection.iSortPos = objectsShop.Ссылка.Родитель.ор_Сезон.Сортировка;

                            ti.group.collectionForSite.strCodeSync = objectsShop.Ссылка.Родитель.op_СезонДляСайта.КодДляСинхронизации;
                            ti.group.collectionForSite.strName = objectsShop.Ссылка.Родитель.op_СезонДляСайта.Наименование;
                            ti.group.collectionForSite.bInvisible = objectsShop.Ссылка.Родитель.op_СезонДляСайта.Невидимость;
                            ti.group.collectionForSite.iSortPos = objectsShop.Ссылка.Родитель.op_СезонДляСайта.Сортировка;

                            ti.type.strCodeSync = objectsShop.Ссылка.ВидНоменклатуры.КодДляСинхронизации;
                            ti.type.strName = objectsShop.Ссылка.ВидНоменклатуры.Наименование;
                            ti.type.bInvisible = objectsShop.Ссылка.ВидНоменклатуры.Невидимость;

                            listTovarsShop.Add(ti);

                            backgroundWorker3.ReportProgress(-1, "Получение данных по номенклатурам из базы магазина (" + listTovarsShop.Count.ToString() + ").");
                        }

                        Marshal.FinalReleaseComObject(objectsShop2);
                        Marshal.FinalReleaseComObject(objectsShop);
                        Marshal.FinalReleaseComObject(queryShop);
                    }

                    listBrands.Clear();
                    listBrandsShop.Clear();

                    // бренды из основной базы
                    {
                        backgroundWorker3.ReportProgress(0, "Получение данных по брендам из основной базы.");

                        // номенклатуры из основной базы
                        dynamic query = connect.NewObject("Запрос");
                        query.Текст = "ВЫБРАТЬ Код, КодДляСинхронизации, Ссылка, ДатаИзменения ИЗ Справочник.ор_ТорговыеМарки ГДЕ ДатаИзменения > &Data";

                        query.УстановитьПараметр("Data", DateSync);

                        dynamic objects2 = query.Выполнить();
                        dynamic objects = objects2.Выбрать();

                        BrandInfo bi;
                        while (objects.Следующий())
                        {
                            bi = new BrandInfo();
                            //bi.ref_ = objects.Ссылка;
                            bi.strCodeSync = objects.КодДляСинхронизации;
                            bi.dateChange = objects.ДатаИзменения;
                            bi.strName = objects.Ссылка.Наименование;
                            bi.iSortPos = objects.Ссылка.Сортировка;
                            bi.bDeleted = objects.Ссылка.ПометкаУдаления;

                            // добавляем товары в документ
                            dynamic sizes = objects.Ссылка.СоответствиеРазмеров.Выгрузить();
                            int count = sizes.Количество();

                            SizeInfo si;
                            for (int i = 0; i < count; i++)
                            {
                                dynamic row = sizes.Получить(i);
                                si = new SizeInfo();
                                si.strSize = row.Размер.Наименование;
                                si.iRussianSize = row.РусскийРазмер;

                                bi.listSizes.Add(si);

                                Marshal.FinalReleaseComObject(row);
                            }

                            Marshal.FinalReleaseComObject(sizes);

                            listBrands.Add(bi);

                            backgroundWorker3.ReportProgress(-1, "Получение данных по брендам из основной базы (" + listBrands.Count.ToString() + ").");
                        }

                        Marshal.FinalReleaseComObject(objects2);
                        Marshal.FinalReleaseComObject(objects);
                        Marshal.FinalReleaseComObject(query);
                    }

                    // бренды из базы магазина
                    {
                        backgroundWorker3.ReportProgress(0, "Получение данных по брендам из базы магазина.");

                        // номенклатуры из основной базы
                        dynamic queryShop = connectShop.NewObject("Запрос");
                        queryShop.Текст = "ВЫБРАТЬ Код, КодДляСинхронизации, Ссылка, ДатаИзменения ИЗ Справочник.ор_ТорговыеМарки ГДЕ ДатаИзменения > &Data";

                        queryShop.УстановитьПараметр("Data", dateSyncShop);

                        dynamic objects2 = queryShop.Выполнить();
                        dynamic objects = objects2.Выбрать();

                        BrandInfo bi;
                        while (objects.Следующий())
                        {
                            bi = new BrandInfo();
                            //bi.ref_ = objects.Ссылка;
                            bi.strCodeSync = objects.КодДляСинхронизации;
                            bi.dateChange = objects.ДатаИзменения;
                            bi.strName = objects.Ссылка.Наименование;
                            bi.iSortPos = objects.Ссылка.Сортировка;
                            bi.bDeleted = objects.Ссылка.ПометкаУдаления;

                            // добавляем товары в документ
                            dynamic sizes = objects.Ссылка.СоответствиеРазмеров.Выгрузить();
                            int count = sizes.Количество();

                            SizeInfo si;
                            for (int i = 0; i < count; i++)
                            {
                                dynamic row = sizes.Получить(i);
                                si = new SizeInfo();
                                si.strSize = row.Размер.Наименование;
                                si.iRussianSize = row.РусскийРазмер;

                                bi.listSizes.Add(si);

                                Marshal.FinalReleaseComObject(row);
                            }

                            Marshal.FinalReleaseComObject(sizes);

                            listBrandsShop.Add(bi);

                            backgroundWorker3.ReportProgress(-1, "Получение данных по брендам из базы магазина (" + listBrandsShop.Count.ToString() + ").");
                        }

                        Marshal.FinalReleaseComObject(objects2);
                        Marshal.FinalReleaseComObject(objects);
                        Marshal.FinalReleaseComObject(queryShop);
                    }
                }

                //if (bTest) return;

                // -----------------------
                // документы установки цен
                // -----------------------
                {
                    Cursor.Current = Cursors.WaitCursor;

                    {
                        listPriceDocuments.Clear();

                        backgroundWorker3.ReportProgress(0, "Получение данных по документам установки цены из основной базы.");

                        dynamic queryPD = connect.NewObject("Запрос");
                        queryPD.Текст = "ВЫБРАТЬ Номер, Ссылка, Дата, Комментарий, Проведен, ПометкаУдаления, КодДляСинхронизации ИЗ Документ.УстановкаЦенНоменклатуры ГДЕ ДатаИзменения > &Data";

                        queryPD.УстановитьПараметр("Data", DateSync);

                        //MessageBox.Show(DateSync.ToShortDateString() + DateSync.ToShortTimeString());

                        dynamic objectsPD2 = queryPD.Выполнить();
                        dynamic objectsPD = objectsPD2.Выбрать();

                        PriceDocumentInfo pdi;
                        while (objectsPD.Следующий())
                        {
                            pdi = new PriceDocumentInfo();
                            //pdi.ref_ = objectsPD.Ссылка;
                            pdi.strCodeSync = objectsPD.КодДляСинхронизации;
                            pdi.strCode = objectsPD.Номер;
                            pdi.date = objectsPD.Дата;
                            pdi.strNote = objectsPD.Комментарий;
                            pdi.bProvodka = objectsPD.Проведен;
                            pdi.bDeleted = objectsPD.ПометкаУдаления;

                            // добавляем товары в документ
                            dynamic tovars = objectsPD.Ссылка.Товары.Выгрузить();//,"КлючСвязи, Номенклатура, Характеристика, Цена");
                            int count = tovars.Количество();

                            TovarPriceInfo tpi;
                            for (int i = 0; i < count; i++)
                            {
                                dynamic row = tovars.Получить(i);
                                tpi = new TovarPriceInfo();
                                tpi.strTovarSyncCode = row.Номенклатура.КодДляСинхронизации;
                                tpi.strAttribSyncCode = row.Характеристика.КодДляСинхронизации;
                                tpi.price = row.Цена;

                                pdi.listTovars.Add(tpi);

                                Marshal.FinalReleaseComObject(row);
                            }

                            Marshal.FinalReleaseComObject(tovars);

                            listPriceDocuments.Add(pdi);
                            backgroundWorker3.ReportProgress(-1, "Получение данных по документам установки цены из основной базы (" + listPriceDocuments.Count.ToString() + ").");
                        }

                        Marshal.FinalReleaseComObject(objectsPD2);
                        Marshal.FinalReleaseComObject(objectsPD);
                        Marshal.FinalReleaseComObject(queryPD);
                    }

                    {
                        backgroundWorker3.ReportProgress(0, "Получение данных по документам установки цены из базы магазина.");
                    
                        listPriceDocumentsShop.Clear();

                        dynamic queryPD = connectShop.NewObject("Запрос");
                        queryPD.Текст = "ВЫБРАТЬ Номер, Ссылка, Дата, Комментарий, Проведен, ПометкаУдаления, КодДляСинхронизации ИЗ Документ.УстановкаЦенНоменклатуры ГДЕ ДатаИзменения > &Data";

                        //DateTime dateSyncShop = connectShop.Константы.ДатаСинхронизации.Получить();
                        queryPD.УстановитьПараметр("Data", dateSyncShop);

                        //MessageBox.Show(DateSync.ToShortDateString() + DateSync.ToShortTimeString());

                        dynamic objectsPD2 = queryPD.Выполнить();
                        dynamic objectsPD = objectsPD2.Выбрать();

                        PriceDocumentInfo pdi;
                        while (objectsPD.Следующий())
                        {
                            pdi = new PriceDocumentInfo();
                            //pdi.ref_ = objectsPD.Ссылка;
                            pdi.strCodeSync = objectsPD.КодДляСинхронизации;
                            pdi.strCode = objectsPD.Номер;
                            pdi.date = objectsPD.Дата;
                            pdi.strNote = objectsPD.Комментарий;
                            pdi.bProvodka = objectsPD.Проведен;
                            pdi.bDeleted = objectsPD.ПометкаУдаления;

                            // добавляем товары в документ
                            dynamic tovars = objectsPD.Ссылка.Товары.Выгрузить();//,"КлючСвязи, Номенклатура, Характеристика, Цена");
                            int count = tovars.Количество();

                            TovarPriceInfo tpi;
                            for (int i = 0; i < count; i++)
                            {
                                dynamic row = tovars.Получить(i);
                                tpi = new TovarPriceInfo();
                                tpi.strTovarSyncCode = row.Номенклатура.КодДляСинхронизации;
                                tpi.strAttribSyncCode = row.Характеристика.КодДляСинхронизации;
                                tpi.price = row.Цена;

                                pdi.listTovars.Add(tpi);

                                Marshal.FinalReleaseComObject(row);
                            }

                            Marshal.FinalReleaseComObject(tovars);

                            listPriceDocumentsShop.Add(pdi);
                            backgroundWorker3.ReportProgress(-1, "Получение данных по документам установки цены из базы магазина (" + listPriceDocumentsShop.Count.ToString() + ").");
                        }

                        Marshal.FinalReleaseComObject(objectsPD2);
                        Marshal.FinalReleaseComObject(objectsPD);
                        Marshal.FinalReleaseComObject(queryPD);
                    }
                }

                // ---------------------
                // Маркетинговые акции и дисконтные карты
                // ---------------------
                {
                    Cursor.Current = Cursors.WaitCursor;

                    backgroundWorker3.ReportProgress(0, "Получение данных по скидкам из основной базы.");
                    // получаем данные на сервере
                    {
                        if (ActionDocument != null && ActionDocument.ref_ != null) Marshal.FinalReleaseComObject(ActionDocument.ref_);

                        ActionDocument = null;

                        dynamic actions = connect.Документы.МаркетинговаяАкция.Выбрать();
                        while (actions.Следующий())
                        {
                            ActionDocument = new ActionInfo();

                            ActionDocument.bDeleted = actions.ПометкаУдаления;
                            ActionDocument.bProvodka = actions.Проведен;
                            ActionDocument.dateBegin = actions.ДатаНачалаДействия;
                            ActionDocument.dateEnd = actions.ДатаОкончанияДействия;
                            ActionDocument.strName = actions.НаименованиеАкции;
                            ActionDocument.ref_ = actions.ПолучитьОбъект();

                            // добавляем скидки в документ
                            dynamic discounts = actions.Ссылка.СкидкиНаценки.Выгрузить();//,"КлючСвязи, Номенклатура, Характеристика, Цена");
                            int count = discounts.Количество();

                            DiscountInfo di;
                            for (int i = 0; i < count; i++)
                            {
                                dynamic row = discounts.Получить(i);
                                di = new DiscountInfo();
                                di.strName = row.СкидкаНаценка.Наименование;
                                di.fProcent = row.СкидкаНаценка.ЗначениеСкидкиНаценки;

                                ActionDocument.listDiscounts.Add(di);

                                Marshal.FinalReleaseComObject(row);
                            }

                            Marshal.FinalReleaseComObject(discounts);

                            break;
                        }

                        backgroundWorker3.ReportProgress(-1, "Получение данных по скидкам из основной базы (ok).");

                        Marshal.FinalReleaseComObject(actions);
                    }

                    backgroundWorker3.ReportProgress(0, "Получение данных по скидкам из базы магазина.");
                    // получаем данные в магазине
                    {
                        Cursor.Current = Cursors.WaitCursor;

                        if (ActionShopDocument != null && ActionShopDocument.ref_ != null) Marshal.FinalReleaseComObject(ActionShopDocument.ref_);

                        ActionShopDocument = null;

                        dynamic actions = connectShop.Документы.МаркетинговаяАкция.Выбрать();
                        while (actions.Следующий())
                        {
                            ActionShopDocument = new ActionInfo();

                            ActionShopDocument.bDeleted = actions.ПометкаУдаления;
                            ActionShopDocument.bProvodka = actions.Проведен;
                            ActionShopDocument.dateBegin = actions.ДатаНачалаДействия;
                            ActionShopDocument.dateEnd = actions.ДатаОкончанияДействия;
                            ActionShopDocument.strName = actions.НаименованиеАкции;
                            ActionShopDocument.ref_ = actions.ПолучитьОбъект();

                            // добавляем скидки в документ
                            dynamic discounts = actions.Ссылка.СкидкиНаценки.Выгрузить();
                            int count = discounts.Количество();

                            DiscountInfo di;
                            for (int i = 0; i < count; i++)
                            {
                                dynamic row = discounts.Получить(i);
                                di = new DiscountInfo();
                                di.strName = row.СкидкаНаценка.Наименование;
                                di.fProcent = row.СкидкаНаценка.ЗначениеСкидкиНаценки;

                                ActionShopDocument.listDiscounts.Add(di);

                                Marshal.FinalReleaseComObject(row);
                            }

                            Marshal.FinalReleaseComObject(discounts);

                            break;
                        }

                        backgroundWorker3.ReportProgress(-1, "Получение данных по скидкам из базы магазина (ok).");

                        Marshal.FinalReleaseComObject(actions);
                    }
                }


                // -------
                // контрагенты
                // -------
                {
                    Cursor.Current = Cursors.WaitCursor;

                    backgroundWorker3.ReportProgress(0, "Получение данных по контрагентам из основной базы.");
                    // в основной базе
                    {
                        listSuppliers.Clear();

                        dynamic query = connect.NewObject("Запрос");
                        query.Текст = "ВЫБРАТЬ Код, Ссылка ИЗ Справочник.Контрагенты ГДЕ ДатаИзменения > &Data";

                        query.УстановитьПараметр("Data", DateSync);

                        dynamic objects2 = query.Выполнить();
                        dynamic objects = objects2.Выбрать();

                        SupplierInfo si;
                        while (objects.Следующий())
                        {
                            si = new SupplierInfo();
                            si.strCodeSync = objects.Ссылка.КодДляСинхронизации;
                            si.strName = objects.Ссылка.Наименование;
                            si.strFullName = objects.Ссылка.НаименованиеПолное;
                            si.strINN = objects.Ссылка.ИНН;
                            si.strNote = objects.Ссылка.Комментарий;

                            si.strSurname = objects.Ссылка.Фамилия;
                            si.strFirstname = objects.Ссылка.Имя;
                            si.strFathername = objects.Ссылка.Отчество;

                            si.bWholeSaler = objects.Ссылка.Оптовик;
                            si.bNoAgreeWithDelivery = objects.Ссылка.ОтказОтСМСРассылки;
                            si.bIgnor = objects.Ссылка.Игнор;

                            si.strBrandWishes = objects.Ссылка.ИнтересующиеМарки;
                            si.strSizeWishes = objects.Ссылка.ИнтересующиеРазмеры;
                            si.strCategoryWishes = objects.Ссылка.ИнтересующиеКатегории;

                            si.strWishes1 = "";
                            si.strWishes2 = "";
                            si.strWishes3 = "";

                            if (objects.Ссылка.Предпочтения1 != connect.Справочники.КонтрагентыПредпочтения.ПустаяСсылка())
                                si.strWishes1 = objects.Ссылка.Предпочтения1.Наименование;
                            if (objects.Ссылка.Предпочтения2 != connect.Справочники.КонтрагентыПредпочтения.ПустаяСсылка())
                                si.strWishes2 = objects.Ссылка.Предпочтения2.Наименование;
                            if (objects.Ссылка.Предпочтения3 != connect.Справочники.КонтрагентыПредпочтения.ПустаяСсылка())
                                si.strWishes3 = objects.Ссылка.Предпочтения3.Наименование;

                            si.dateChange = objects.Ссылка.ДатаИзменения;
                            dynamic val = objects.Ссылка.ЮрФизЛицо;
                            if (connect.String(val) == connect.String(connect.Перечисления.ЮрФизЛицо.ЮрЛицо))
                                si.Type = 0;
                            else
                                si.Type = 1;

                            dynamic ex_info = objects.Ссылка.КонтактнаяИнформация.Выгрузить();
                            int count = ex_info.Количество();

                            for (int i = 0; i < count; i++)
                            {
                                dynamic row = ex_info.Получить(i);
                                si.dictExInfo[row.Вид.Наименование] = row.Представление;
                                Marshal.FinalReleaseComObject(row);
                            }

                            listSuppliers.Add(si);

                            backgroundWorker3.ReportProgress(-1, "Получение данных по контрагентам из основной базы (" + listSuppliers.Count.ToString() + ").");

                            Marshal.FinalReleaseComObject(val);
                        }

                        Marshal.FinalReleaseComObject(objects2);
                        Marshal.FinalReleaseComObject(objects);
                        Marshal.FinalReleaseComObject(query);
                    }

                    backgroundWorker3.ReportProgress(0, "Получение данных по контрагентам из базы магазина.");
                    // в базе магазина
                    {
                        listSuppliersShop.Clear();

                        dynamic query = connectShop.NewObject("Запрос");
                        query.Текст = "ВЫБРАТЬ Код, Ссылка ИЗ Справочник.Контрагенты ГДЕ ДатаИзменения > &Data";

                        //DateTime dateSyncShop = connectShop.Константы.ДатаСинхронизации.Получить();

                        query.УстановитьПараметр("Data", dateSyncShop);

                        dynamic objects2 = query.Выполнить();
                        dynamic objects = objects2.Выбрать();

                        SupplierInfo si;
                        while (objects.Следующий())
                        {
                            si = new SupplierInfo();
                            si.strCodeSync = objects.Ссылка.КодДляСинхронизации;
                            si.strName = objects.Ссылка.Наименование;
                            si.strFullName = objects.Ссылка.НаименованиеПолное;
                            si.strINN = objects.Ссылка.ИНН;
                            si.strNote = objects.Ссылка.Комментарий;

                            si.strSurname = objects.Ссылка.Фамилия;
                            si.strFirstname = objects.Ссылка.Имя;
                            si.strFathername = objects.Ссылка.Отчество;

                            si.bWholeSaler = objects.Ссылка.Оптовик;
                            si.bNoAgreeWithDelivery = objects.Ссылка.ОтказОтСМСРассылки;
                            si.bIgnor = objects.Ссылка.Игнор;

                            si.strBrandWishes = objects.Ссылка.ИнтересующиеМарки;
                            si.strSizeWishes = objects.Ссылка.ИнтересующиеРазмеры;
                            si.strCategoryWishes = objects.Ссылка.ИнтересующиеКатегории;

                            si.strWishes1 = "";
                            si.strWishes2 = "";
                            si.strWishes3 = "";

                            if (objects.Ссылка.Предпочтения1 != connectShop.Справочники.КонтрагентыПредпочтения.ПустаяСсылка())
                                si.strWishes1 = objects.Ссылка.Предпочтения1.Наименование;
                            if (objects.Ссылка.Предпочтения2 != connectShop.Справочники.КонтрагентыПредпочтения.ПустаяСсылка())
                                si.strWishes2 = objects.Ссылка.Предпочтения2.Наименование;
                            if (objects.Ссылка.Предпочтения3 != connectShop.Справочники.КонтрагентыПредпочтения.ПустаяСсылка())
                                si.strWishes3 = objects.Ссылка.Предпочтения3.Наименование;

                            si.dateChange = objects.Ссылка.ДатаИзменения;
                            dynamic val = objects.Ссылка.ЮрФизЛицо;
                            if (connectShop.String(val) == connectShop.String(connectShop.Перечисления.ЮрФизЛицо.ЮрЛицо))
                                si.Type = 0;
                            else
                                si.Type = 1;

                            dynamic ex_info = objects.Ссылка.КонтактнаяИнформация.Выгрузить();
                            int count = ex_info.Количество();

                            for (int i = 0; i < count; i++)
                            {
                                dynamic row = ex_info.Получить(i);
                                si.dictExInfo[row.Вид.Наименование] = row.Представление;
                                Marshal.FinalReleaseComObject(row);
                            }

                            listSuppliersShop.Add(si);

                            backgroundWorker3.ReportProgress(-1, "Получение данных по контрагентам из базы магазина (" + listSuppliersShop.Count.ToString() + ").");

                            Marshal.FinalReleaseComObject(val);
                        }

                        Marshal.FinalReleaseComObject(objects2);
                        Marshal.FinalReleaseComObject(objects);
                        Marshal.FinalReleaseComObject(query);
                    }

                }

                // ---------------------
                // Документы прихода
                // ---------------------
                {
                    Cursor.Current = Cursors.WaitCursor;
                
                    backgroundWorker3.ReportProgress(0, "Получение данных по документам прихода.");
                    // в базе магазина
                    {
                        listInvoiceInDocumentsShop.Clear();

                        dynamic queryIn = connectShop.NewObject("Запрос");
                        queryIn.Текст = "ВЫБРАТЬ Номер, Ссылка, Дата, Проведен, ПометкаУдаления, КодДляСинхронизации, ДатаИзменения, Комментарий, Отчет, Ответственный ИЗ Документ.ПоступлениеТоваров " +
                            "ГДЕ ДатаИзменения > &Data";

                        //DateTime dateSyncDocumentsIn = connectShop.Константы.ДатаСинхронизации.Получить();

                        queryIn.УстановитьПараметр("Data", dateSyncShop);

                        dynamic objectsIn2 = queryIn.Выполнить();
                        dynamic objectsIn = objectsIn2.Выбрать();

                        InvoiceInDocumentInfo iidi;
                        while (objectsIn.Следующий())
                        {
                            iidi = new InvoiceInDocumentInfo();
                            iidi.strCode = objectsIn.Номер;
                            iidi.date = objectsIn.Дата;
                            iidi.strCodeSync = objectsIn.КодДляСинхронизации;
                            iidi.dateChange = objectsIn.ДатаИзменения;
                            iidi.bProvodka = objectsIn.Проведен;
                            iidi.bDeleted = objectsIn.ПометкаУдаления;
                            iidi.bReport = objectsIn.Отчет;
                            iidi.strNote = objectsIn.Комментарий;
                            iidi.strUser = objectsIn.Ответственный.Наименование;

                            iidi.strSupplierCodeSync = objectsIn.Ссылка.Контрагент.КодДляСинхронизации;

                            // добавляем товары в документ
                            dynamic tovars = objectsIn.Ссылка.Товары.Выгрузить();//,"КлючСвязи, Номенклатура, Характеристика, Цена");
                            int count = tovars.Количество();

                            TovarInvoiceInfo tpi;
                            for (int i = 0; i < count; i++)
                            {
                                dynamic row = tovars.Получить(i);
                                tpi = new TovarInvoiceInfo();
                                tpi.strTovarSyncCode = row.Номенклатура.КодДляСинхронизации;
                                tpi.strAttribSyncCode = row.Характеристика.КодДляСинхронизации;
                                tpi.price = row.Цена;
                                tpi.iCount = row.Количество;

                                iidi.listTovars.Add(tpi);

                                Marshal.FinalReleaseComObject(row);
                            }

                            Marshal.FinalReleaseComObject(tovars);

                            listInvoiceInDocumentsShop.Add(iidi);

                            backgroundWorker3.ReportProgress(-1, "Получение данных по документам прихода (" + listInvoiceInDocumentsShop.Count.ToString() + ").");
                        }

                        Marshal.FinalReleaseComObject(objectsIn2);
                        Marshal.FinalReleaseComObject(objectsIn);
                        Marshal.FinalReleaseComObject(queryIn);
                    }
                }


                // ---------------------
                // Документы реализации
                // ---------------------
                {
                    Cursor.Current = Cursors.WaitCursor;

                    backgroundWorker3.ReportProgress(0, "Получение данных по документам расхода.");
                    // в базе магазина
                    {
                        /*for (int i = 0; i < listInvoiceOutDocumentsShop.Count; i++)
                        {
                            if (listInvoiceOutDocumentsShop[i].ref_ != null) Marshal.FinalReleaseComObject(listInvoiceOutDocumentsShop[i].ref_);
                        }*/
                        listInvoiceOutDocumentsShop.Clear();

                        dynamic queryOut = connectShop.NewObject("Запрос");
                        queryOut.Текст = "ВЫБРАТЬ Номер, Ссылка, Дата, СуммаДокумента, Проведен, ПометкаУдаления, Отчет, " + 
                            "КодДляСинхронизации, Комментарий, ДатаИзменения, Ответственный, ОтправкаПоДоговору, Оптовик ИЗ Документ.РеализацияТоваров " +
                            " ГДЕ ДатаИзменения > &Data";

                        //DateTime dateSyncDocumentsOut = connectShop.Константы.ДатаСинхронизации.Получить();

                        queryOut.УстановитьПараметр("Data", dateSyncShop);

                        dynamic objectsOut2 = queryOut.Выполнить();
                        dynamic objectsOut = objectsOut2.Выбрать();

                        InvoiceOutDocumentInfo iodi;
                        while (objectsOut.Следующий())
                        {
                            iodi = new InvoiceOutDocumentInfo();
                            //iodi.ref_ = objectsOut.Ссылка;
                            iodi.strCodeSync = objectsOut.КодДляСинхронизации;
                            iodi.strCode = objectsOut.Номер;
                            iodi.date = objectsOut.Дата;
                            iodi.dateChange = objectsOut.ДатаИзменения;
                            iodi.bProvodka = objectsOut.Проведен;
                            iodi.bDeleted = objectsOut.ПометкаУдаления;
                            iodi.fDocumentSum = objectsOut.СуммаДокумента;
                            iodi.bReport = objectsOut.Отчет;
                            iodi.bOptovik = objectsOut.Оптовик;
                            iodi.strNote = objectsOut.Комментарий;
                            iodi.strDogovor = objectsOut.ОтправкаПоДоговору;
                            iodi.strUser = objectsOut.Ответственный.Наименование;

                            if (objectsOut.Ссылка.ДисконтнаяКарта != connectShop.Справочники.ИнформационныеКарты.ПустаяСсылка())
                            {
                                iodi.card.strName = objectsOut.Ссылка.ДисконтнаяКарта.Наименование;
                                iodi.card.strCode = objectsOut.Ссылка.ДисконтнаяКарта.КодКарты;
                            }
                            else
                            {
                                iodi.card.strName = "";
                                iodi.card.strCode = "";
                            }

                            iodi.strSupplierCodeSync = objectsOut.Ссылка.Контрагент.КодДляСинхронизации;

                            // добавляем товары в документ
                            dynamic tovars = objectsOut.Ссылка.Товары.Выгрузить();
                            int count = tovars.Количество();

                            TovarInvoiceInfo tpi;
                            for (int i = 0; i < count; i++)
                            {
                                dynamic row = tovars.Получить(i);
                                tpi = new TovarInvoiceInfo();
                                tpi.strTovarSyncCode = row.Номенклатура.КодДляСинхронизации;
                                tpi.strAttribSyncCode = row.Характеристика.КодДляСинхронизации;
                                tpi.price = row.Цена;
                                tpi.manual_discount = row.ПроцентРучнойСкидки;
                                tpi.manual_discount_opt = row.ПроцентРучнойСкидкиОпт;
                                tpi.summa_manual_discount = row.СуммаРучнойСкидки;
                                tpi.iCount = row.Количество;

                                iodi.listTovars.Add(tpi);

                                Marshal.FinalReleaseComObject(row);
                            }

                            Marshal.FinalReleaseComObject(tovars);

                            listInvoiceOutDocumentsShop.Add(iodi);

                            backgroundWorker3.ReportProgress(-1, "Получение данных по документам расхода (" + listInvoiceOutDocumentsShop.Count.ToString() + ").");
                        }

                        Marshal.FinalReleaseComObject(objectsOut2);
                        Marshal.FinalReleaseComObject(objectsOut);
                        Marshal.FinalReleaseComObject(queryOut);
                    }
                }

                // ---------------------
                // Документы возврата от покупателя
                // ---------------------
                {
                    Cursor.Current = Cursors.WaitCursor;

                    backgroundWorker3.ReportProgress(0, "Получение данных по документам возврата.");
                    // в базе магазина
                    {
                        listInvoiceReturnDocumentsShop.Clear();

                        dynamic queryReturn = connectShop.NewObject("Запрос");
                        queryReturn.Текст = "ВЫБРАТЬ Номер, Ссылка, Дата, Проведен, ПометкаУдаления, Отчет, КодДляСинхронизации, Комментарий, ДатаИзменения, Ответственный, Оптовик, РучнаяЦена ИЗ Документ.ВозвратТоваровОтПокупателя " +
                            " ГДЕ ДатаИзменения > &Data";

                        //DateTime dateSyncDocuments = connectShop.Константы.ДатаСинхронизации.Получить();

                        queryReturn.УстановитьПараметр("Data", dateSyncShop);

                        dynamic objectsReturn2 = queryReturn.Выполнить();
                        dynamic objectsReturn = objectsReturn2.Выбрать();

                        InvoiceReturnDocumentInfo irdi;
                        while (objectsReturn.Следующий())
                        {
                            irdi = new InvoiceReturnDocumentInfo();
                            //iodi.ref_ = objectsOut.Ссылка;
                            irdi.strCodeSync = objectsReturn.КодДляСинхронизации;
                            irdi.strCode = objectsReturn.Номер;
                            irdi.date = objectsReturn.Дата;
                            irdi.dateChange = objectsReturn.ДатаИзменения;
                            irdi.bProvodka = objectsReturn.Проведен;
                            irdi.bDeleted = objectsReturn.ПометкаУдаления;
                            irdi.strAnaliticName = objectsReturn.Ссылка.АналитикаХозяйственнойОперации.Наименование;
                            irdi.bReport = objectsReturn.Отчет;
                            irdi.bOptovik = objectsReturn.Оптовик;
                            irdi.bManualCost = objectsReturn.РучнаяЦена;
                            irdi.strNote = objectsReturn.Комментарий;
                            irdi.strUser = objectsReturn.Ответственный.Наименование;

                            if (objectsReturn.Ссылка.ДисконтнаяКарта != connectShop.Справочники.ИнформационныеКарты.ПустаяСсылка())
                            {
                                irdi.card.strName = objectsReturn.Ссылка.ДисконтнаяКарта.Наименование;
                                irdi.card.strCode = objectsReturn.Ссылка.ДисконтнаяКарта.КодКарты;
                            }
                            else
                            {
                                irdi.card.strName = "";
                                irdi.card.strCode = "";
                            }

                            irdi.strSupplierCodeSync = objectsReturn.Ссылка.Контрагент.КодДляСинхронизации;

                            // добавляем товары в документ
                            dynamic tovars = objectsReturn.Ссылка.Товары.Выгрузить();
                            int count = tovars.Количество();

                            TovarInvoiceInfo tpi;
                            for (int i = 0; i < count; i++)
                            {
                                dynamic row = tovars.Получить(i);
                                tpi = new TovarInvoiceInfo();
                                tpi.strTovarSyncCode = row.Номенклатура.КодДляСинхронизации;
                                tpi.strAttribSyncCode = row.Характеристика.КодДляСинхронизации;
                                tpi.price = row.Цена;
                                tpi.manual_discount = row.ПроцентСкидки;
                                tpi.iCount = row.Количество;
                                //tpi.strInvoiceInCodeSync = row.ДокументПродажи.КодДляСинхронизации;

                                irdi.listTovars.Add(tpi);

                                Marshal.FinalReleaseComObject(row);
                            }

                            Marshal.FinalReleaseComObject(tovars);

                            listInvoiceReturnDocumentsShop.Add(irdi);

                            backgroundWorker3.ReportProgress(-1, "Получение данных по документам возврата (" + listInvoiceReturnDocumentsShop.Count.ToString() + ").");
                        }

                        Marshal.FinalReleaseComObject(objectsReturn2);
                        Marshal.FinalReleaseComObject(objectsReturn);
                        Marshal.FinalReleaseComObject(queryReturn);
                    }
                }

                // ---------------------
                // Документы оприходования
                // ---------------------
                {
                    Cursor.Current = Cursors.WaitCursor;

                    backgroundWorker3.ReportProgress(0, "Получение данных по документам оприходования.");
                    // в базе магазина
                    {
                        listInvoiceMoveInDocumentsShop.Clear();

                        dynamic queryReturn = connectShop.NewObject("Запрос");
                        queryReturn.Текст = "ВЫБРАТЬ Номер, Ссылка, Дата, Проведен, ПометкаУдаления, КодДляСинхронизации, Комментарий, ДатаИзменения, Ответственный ИЗ Документ.ОприходованиеТоваров " +
                            " ГДЕ ДатаИзменения > &Data И Фиктивный = ложь";

                        //DateTime dateSyncDocuments = connectShop.Константы.ДатаСинхронизации.Получить();

                        queryReturn.УстановитьПараметр("Data", dateSyncShop);

                        dynamic objectsReturn2 = queryReturn.Выполнить();
                        dynamic objectsReturn = objectsReturn2.Выбрать();

                        InvoiceMoveInOutDocumentInfo irdi;
                        while (objectsReturn.Следующий())
                        {
                            irdi = new InvoiceMoveInOutDocumentInfo();
                            //iodi.ref_ = objectsOut.Ссылка;
                            irdi.strCodeSync = objectsReturn.КодДляСинхронизации;
                            irdi.strCode = objectsReturn.Номер;
                            irdi.date = objectsReturn.Дата;
                            irdi.dateChange = objectsReturn.ДатаИзменения;
                            irdi.bProvodka = objectsReturn.Проведен;
                            irdi.bDeleted = objectsReturn.ПометкаУдаления;
                            irdi.strNote = objectsReturn.Комментарий;
                            irdi.strUser = objectsReturn.Ответственный.Наименование;

                            // добавляем товары в документ
                            dynamic tovars = objectsReturn.Ссылка.Товары.Выгрузить();
                            int count = tovars.Количество();

                            TovarInvoiceInfo tpi;
                            for (int i = 0; i < count; i++)
                            {
                                dynamic row = tovars.Получить(i);
                                tpi = new TovarInvoiceInfo();
                                tpi.strTovarSyncCode = row.Номенклатура.КодДляСинхронизации;
                                tpi.strAttribSyncCode = row.Характеристика.КодДляСинхронизации;
                                tpi.price = row.Цена;
                                tpi.iCount = row.Количество;
                                //tpi.strInvoiceInCodeSync = row.ДокументПродажи.КодДляСинхронизации;

                                irdi.listTovars.Add(tpi);

                                Marshal.FinalReleaseComObject(row);
                            }

                            Marshal.FinalReleaseComObject(tovars);

                            listInvoiceMoveInDocumentsShop.Add(irdi);

                            backgroundWorker3.ReportProgress(-1, "Получение данных по документам оприходования (" + listInvoiceMoveInDocumentsShop.Count.ToString() + ").");
                        }

                        Marshal.FinalReleaseComObject(objectsReturn2);
                        Marshal.FinalReleaseComObject(objectsReturn);
                        Marshal.FinalReleaseComObject(queryReturn);
                    }
                }

                // ---------------------
                // Документы списания
                // ---------------------
                {
                    Cursor.Current = Cursors.WaitCursor;

                    backgroundWorker3.ReportProgress(0, "Получение данных по документам списания.");
                    // в базе магазина
                    {
                        listInvoiceMoveOutDocumentsShop.Clear();

                        dynamic queryReturn = connectShop.NewObject("Запрос");
                        queryReturn.Текст = "ВЫБРАТЬ Номер, Ссылка, Дата, Проведен, ПометкаУдаления, КодДляСинхронизации, Комментарий, ДатаИзменения, Ответственный ИЗ Документ.СписаниеТоваров " +
                            " ГДЕ ДатаИзменения > &Data И Фиктивный = ложь";

                        //DateTime dateSyncDocuments = connectShop.Константы.ДатаСинхронизации.Получить();

                        queryReturn.УстановитьПараметр("Data", dateSyncShop);

                        dynamic objectsReturn2 = queryReturn.Выполнить();
                        dynamic objectsReturn = objectsReturn2.Выбрать();

                        InvoiceMoveInOutDocumentInfo irdi;
                        while (objectsReturn.Следующий())
                        {
                            irdi = new InvoiceMoveInOutDocumentInfo();
                            //iodi.ref_ = objectsOut.Ссылка;
                            irdi.strCodeSync = objectsReturn.КодДляСинхронизации;
                            irdi.strCode = objectsReturn.Номер;
                            irdi.date = objectsReturn.Дата;
                            irdi.dateChange = objectsReturn.ДатаИзменения;
                            irdi.bProvodka = objectsReturn.Проведен;
                            irdi.bDeleted = objectsReturn.ПометкаУдаления;
                            irdi.strNote = objectsReturn.Комментарий;
                            irdi.strUser = objectsReturn.Ответственный.Наименование;

                            // добавляем товары в документ
                            dynamic tovars = objectsReturn.Ссылка.Товары.Выгрузить();
                            int count = tovars.Количество();

                            TovarInvoiceInfo tpi;
                            for (int i = 0; i < count; i++)
                            {
                                dynamic row = tovars.Получить(i);
                                tpi = new TovarInvoiceInfo();
                                tpi.strTovarSyncCode = row.Номенклатура.КодДляСинхронизации;
                                tpi.strAttribSyncCode = row.Характеристика.КодДляСинхронизации;
                                tpi.price = row.Цена;
                                tpi.iCount = row.Количество;

                                irdi.listTovars.Add(tpi);

                                Marshal.FinalReleaseComObject(row);
                            }

                            Marshal.FinalReleaseComObject(tovars);

                            listInvoiceMoveOutDocumentsShop.Add(irdi);

                            backgroundWorker3.ReportProgress(-1, "Получение данных по документам списания (" + listInvoiceMoveOutDocumentsShop.Count.ToString() + ").");
                        }

                        Marshal.FinalReleaseComObject(objectsReturn2);
                        Marshal.FinalReleaseComObject(objectsReturn);
                        Marshal.FinalReleaseComObject(queryReturn);
                    }
                }

                Cursor.Current = Cursors.Default;
            }
            catch (Exception ex)
            {
                backgroundWorker3.ReportProgress(1, "Ошибка при синхронизации: " + ex.Message);
                Cursor.Current = Cursors.Default;
                return;
            }
        }

        private void DBSync_Load(object sender, EventArgs e)
        {
            Login frm = new Login();
            if (frm.ShowDialog() != System.Windows.Forms.DialogResult.OK)
            {
                this.Close();
                return;
            }
        }

        private void btnRefreshDataSite_Click(object sender, EventArgs e)
        {
            if (m_bWasShopSelect)
            {
                gbSync.Enabled = false;
            }
            gbMain.Enabled = false;
            gbShop.Enabled = false;
            btnRefreshDataSite.Enabled = false;
            btnYML.Enabled = false;

            listView1.Items.Clear();

            backgroundWorker4.RunWorkerAsync();
        }

        private void backgroundWorker4_DoWork(object sender, DoWorkEventArgs e)
        {
            ExchangeData();
        }

        private void backgroundWorker4_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            if (e.ProgressPercentage == -1)
            {
                ListViewItem vi = listView1.Items[listView1.Items.Count - 1];
                vi.Text = (string)e.UserState;
                listView1.EnsureVisible(listView1.Items.Count - 1);
            }
            else
            {
                ListViewItem vi = listView1.Items.Add((string)e.UserState);
                if ((string)e.UserState != "") vi.ImageIndex = e.ProgressPercentage;
                listView1.EnsureVisible(listView1.Items.Count - 1);
            }
        }

        private void backgroundWorker4_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (listView1.Items.Count > 0)
                listView1.EnsureVisible(listView1.Items.Count - 1);

            gbMain.Enabled = true;
            gbShop.Enabled = true;
            gbSync.Enabled = true;
            if (m_bWasShopSelect)
            {
                gbSync.Enabled = true;
            }
            btnRefreshDataSite.Enabled = true;
            btnYML.Enabled = true;
        }

        private void ExchangeData()
        {
            Cursor.Current = Cursors.WaitCursor;

            try
            {
                // рассчитываем кол-во номенклатур на основной базе для выбранного магазина
                backgroundWorker4.ReportProgress(0, "");


                backgroundWorker4.ReportProgress(0, "Обновление цен и остатков на складе.");
                dynamic mod = connect.ExchangeData;
                int res = mod.ПроверитьОстаткиИЦену();
                backgroundWorker4.ReportProgress(2, "Обновление цен и остатков на складе прошло успешно.");

                backgroundWorker4.ReportProgress(0, "Передача данных на сайт.");
                string strres = mod.ВыполнитьСинхронизацию(true);
                if (strres == "")
                    backgroundWorker4.ReportProgress(2, "Передача данных на сайт прошла успешно.");
                else
                    backgroundWorker4.ReportProgress(1, "Ошибка при передача данных на сайт: " + strres);

                Marshal.FinalReleaseComObject(mod);
            }
            catch (Exception ex)
            {
                backgroundWorker4.ReportProgress(1, "Ошибка при синхронизации: " + ex.Message);
                Cursor.Current = Cursors.Default;
                return;
            }


            Cursor.Current = Cursors.Default;
        }

        private void cbVerifyMainBase_CheckedChanged(object sender, EventArgs e)
        {
            if (cbVerifyMainBase.Checked)
            {
                cbNoDeleteDocuments.Enabled = true;
            }
            else
            {
                cbNoDeleteDocuments.Enabled = false;
                cbNoDeleteDocuments.Checked = false;
            }
        }

        private void bgVerifyTovar_DoWork(object sender, DoWorkEventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;

            try
            {
                // рассчитываем кол-во номенклатур на основной базе для выбранного магазина
                bgVerifyTovar.ReportProgress(0, "");

                bgVerifyTovar.ReportProgress(0, "Получение данных из основной базы.");

                dynamic query = connect.NewObject("Запрос");
                query.Текст = "ВЫБРАТЬ РАЗРЕШЕННЫЕ " +
                "		ХарактеристикиНоменклатуры.Владелец КАК Номенклатура, " +
                "		ХарактеристикиНоменклатуры.Ссылка КАК Характеристика, " +
                "		СУММА(ЕСТЬNULL(Остатки.КоличествоОстаток, 0)) КАК КоличествоОстаток, " +
                " 		СУММА(ЕСТЬNULL(Цены.Цена, 0)) КАК Цена " +
                " 	ИЗ " +
                "		Справочник.ХарактеристикиНоменклатуры КАК ХарактеристикиНоменклатуры " +
                "		{ЛЕВОЕ СОЕДИНЕНИЕ РегистрНакопления.ТоварыНаСкладах.Остатки(, {(Номенклатура).* КАК Номенклатура, (Характеристика).* КАК Характеристика}) КАК Остатки " +
                "			ПО Остатки.Номенклатура = ХарактеристикиНоменклатуры.Владелец И Остатки.Характеристика = ХарактеристикиНоменклатуры.Ссылка И Остатки.Склад.Магазин = &Магазин} " +
                "		{ЛЕВОЕ СОЕДИНЕНИЕ РегистрСведений.ЦеныНоменклатуры.СрезПоследних(, {(Номенклатура).* КАК Номенклатура, (Характеристика).* КАК Характеристика}) КАК Цены " +
                "			ПО Цены.Номенклатура = ХарактеристикиНоменклатуры.Владелец И Цены.Характеристика = ХарактеристикиНоменклатуры.Ссылка} " +
                " СГРУППИРОВАТЬ ПО " +
                " 		ХарактеристикиНоменклатуры.Владелец, " +
                " 		ХарактеристикиНоменклатуры.Ссылка";

                query.УстановитьПараметр("Магазин", ref_shop);
                dynamic tovars = query.Выполнить().Выбрать();

                dictPairTovarAttribInfo.Clear();

                int i = 0;
                int cnt = tovars.Количество();
                while (tovars.Следующий())
                {
                    i++;
                    string iTovarCode = tovars.Номенклатура.КодДляСинхронизации;
                    string iAttribCode = tovars.Характеристика.КодДляСинхронизации;

                    PairTovarAttribInfo tai = new PairTovarAttribInfo();
                    tai.strTovarName = tovars.Номенклатура.Наименование;
                    tai.bTovarDeleted = tovars.Номенклатура.ПометкаУдаления;
                    tai.strAttribName = tovars.Характеристика.Наименование;
                    tai.bAttribDeleted = tovars.Характеристика.ПометкаУдаления;
                    tai.iRemain = tovars.КоличествоОстаток;
                    tai.fCost = tovars.Цена;
                    tai.iTovarCode = iTovarCode;
                    tai.iAttribCode = iAttribCode;

                    dictPairTovarAttribInfo[iTovarCode + " " + iAttribCode] = tai;

                    bgVerifyTovar.ReportProgress(-1, "Получение данных из основной базы (" + i.ToString() + " из " + cnt.ToString() + ").");
                }

                if (tovars != null) Marshal.FinalReleaseComObject(tovars);
                if (query != null) Marshal.FinalReleaseComObject(query);

                // -------------------------
                bgVerifyTovar.ReportProgress(0, "Получение данных из базы магазина.");

                dynamic queryShop = connectShop.NewObject("Запрос");
                queryShop.Текст = "ВЫБРАТЬ РАЗРЕШЕННЫЕ " +
                "		ХарактеристикиНоменклатуры.Владелец КАК Номенклатура, " +
                "		ХарактеристикиНоменклатуры.Ссылка КАК Характеристика, " +
                "		СУММА(ЕСТЬNULL(Остатки.КоличествоОстаток, 0)) КАК КоличествоОстаток, " +
                " 		СУММА(ЕСТЬNULL(Цены.Цена, 0)) КАК Цена " +
                " 	ИЗ " +
                "		Справочник.ХарактеристикиНоменклатуры КАК ХарактеристикиНоменклатуры " +
                "		{ЛЕВОЕ СОЕДИНЕНИЕ РегистрНакопления.ТоварыНаСкладах.Остатки(, {(Номенклатура).* КАК Номенклатура, (Характеристика).* КАК Характеристика}) КАК Остатки " +
                "			ПО Остатки.Номенклатура = ХарактеристикиНоменклатуры.Владелец И Остатки.Характеристика = ХарактеристикиНоменклатуры.Ссылка} " +
                "		{ЛЕВОЕ СОЕДИНЕНИЕ РегистрСведений.ЦеныНоменклатуры.СрезПоследних(, {(Номенклатура).* КАК Номенклатура, (Характеристика).* КАК Характеристика}) КАК Цены " +
                "			ПО Цены.Номенклатура = ХарактеристикиНоменклатуры.Владелец И Цены.Характеристика = ХарактеристикиНоменклатуры.Ссылка} " +
                " СГРУППИРОВАТЬ ПО " +
                " 		ХарактеристикиНоменклатуры.Владелец, " +
                " 		ХарактеристикиНоменклатуры.Ссылка";

                dynamic tovarsShop = queryShop.Выполнить().Выбрать();

                dictPairTovarAttribInfoShop.Clear();

                i = 0;
                cnt = tovarsShop.Количество();
                while (tovarsShop.Следующий())
                {
                    i++;
                    string iTovarCode = tovarsShop.Номенклатура.КодДляСинхронизации;
                    string iAttribCode = tovarsShop.Характеристика.КодДляСинхронизации;

                    PairTovarAttribInfo tai = new PairTovarAttribInfo();
                    tai.strTovarName = tovarsShop.Номенклатура.Наименование;
                    tai.bTovarDeleted = tovarsShop.Номенклатура.ПометкаУдаления;
                    tai.strAttribName = tovarsShop.Характеристика.Наименование;
                    tai.bAttribDeleted = tovarsShop.Характеристика.ПометкаУдаления;
                    tai.iRemain = tovarsShop.КоличествоОстаток;
                    tai.fCost = tovarsShop.Цена;
                    tai.iTovarCode = iTovarCode;
                    tai.iAttribCode = iAttribCode;

                    dictPairTovarAttribInfoShop[iTovarCode + " " + iAttribCode] = tai;

                    bgVerifyTovar.ReportProgress(-1, "Получение данных из базы магазина (" + i.ToString() + " из " + cnt.ToString() + ").");
                }

                if (tovarsShop != null) Marshal.FinalReleaseComObject(tovarsShop);
                if (queryShop != null) Marshal.FinalReleaseComObject(queryShop);

                bgVerifyTovar.ReportProgress(0, "Проверка номенклатур и характеристик.");

                // ищем несовпадения
                foreach (KeyValuePair<string, PairTovarAttribInfo> val in dictPairTovarAttribInfo)
                {
                    PairTovarAttribInfo valShop;
                    if (dictPairTovarAttribInfoShop.TryGetValue(val.Key, out valShop))
                    {
                        if (val.Value.bAttribDeleted != valShop.bAttribDeleted)
                        {
                            bgVerifyTovar.ReportProgress(1, "Товар: " + valShop.strTovarName + ", " + valShop.strAttribName + ": различается пометка удаления в характеристике.");

                            // нужно поставить отметки о необходимости синхронизации
                            dynamic find = connectShop.Справочники.Номенклатура.НайтиПоРеквизиту("КодДляСинхронизации", valShop.iTovarCode);
                            dynamic tovar_ = find.ПолучитьОбъект();

                            tovar_.ДатаИзменения = DateTime.Now;
                            tovar_.ДатаИзмененияАттрибутов = DateTime.Now;

                            tovar_.Записать();

                            Marshal.FinalReleaseComObject(find);
                            Marshal.FinalReleaseComObject(tovar_);
                        }
                        if (val.Value.bTovarDeleted != valShop.bTovarDeleted)
                        {
                            bgVerifyTovar.ReportProgress(1, "Товар: " + valShop.strTovarName + ", " + valShop.strAttribName + ": различается пометка удаления в номенклатуре.");

                            // нужно поставить отметки о необходимости синхронизации
                            dynamic find = connectShop.Справочники.Номенклатура.НайтиПоРеквизиту("КодДляСинхронизации", valShop.iTovarCode);
                            dynamic tovar_ = find.ПолучитьОбъект();

                            tovar_.ДатаИзменения = DateTime.Now;

                            tovar_.Записать();

                            Marshal.FinalReleaseComObject(find);
                            Marshal.FinalReleaseComObject(tovar_);
                        }
                        if (val.Value.strAttribName != valShop.strAttribName)
                        {
                            bgVerifyTovar.ReportProgress(1, "Товар: " + valShop.strTovarName + ", " + valShop.strAttribName + ": различается наименование характеристики (" +
                                val.Value.strAttribName + " --- " + valShop.strAttribName + ").");

                            // нужно поставить отметки о необходимости синхронизации
                            dynamic find = connectShop.Справочники.Номенклатура.НайтиПоРеквизиту("КодДляСинхронизации", valShop.iTovarCode);
                            dynamic tovar_ = find.ПолучитьОбъект();

                            tovar_.ДатаИзменения = DateTime.Now;
                            tovar_.ДатаИзмененияАттрибутов = DateTime.Now;

                            tovar_.Записать();

                            Marshal.FinalReleaseComObject(find);
                            Marshal.FinalReleaseComObject(tovar_);

                        }
                        if (val.Value.strTovarName != valShop.strTovarName)
                        {
                            bgVerifyTovar.ReportProgress(1, "Товар: " + valShop.strTovarName + ", " + valShop.strAttribName + ": различается наименование номенклатуры (" +
                                val.Value.strTovarName + " --- " + valShop.strTovarName + ").");

                            // нужно поставить отметки о необходимости синхронизации
                            dynamic find = connectShop.Справочники.Номенклатура.НайтиПоРеквизиту("КодДляСинхронизации", valShop.iTovarCode);
                            dynamic tovar_ = find.ПолучитьОбъект();

                            tovar_.ДатаИзменения = DateTime.Now;

                            tovar_.Записать();

                            Marshal.FinalReleaseComObject(find);
                            Marshal.FinalReleaseComObject(tovar_);

                        }
                        if (val.Value.iRemain != valShop.iRemain)
                            bgVerifyTovar.ReportProgress(1, "Товар: " + valShop.strTovarName + ", " + valShop.strAttribName + ": различается остаток номенклатуры (" +
                            val.Value.iRemain.ToString() + " --- " + valShop.iRemain.ToString() + ").");
                        if (val.Value.fCost != valShop.fCost)
                            bgVerifyTovar.ReportProgress(1, "Товар: " + valShop.strTovarName + ", " + valShop.strAttribName + ": различается цена номенклатуры (" +
                            val.Value.fCost.ToString() + " --- " + valShop.fCost.ToString() + ").");
                    }
                    else
                    {
                        if (!val.Value.bTovarDeleted || !val.Value.bAttribDeleted)
                        {
                            bgVerifyTovar.ReportProgress(1, "Товар " + val.Value.strTovarName + ", " + val.Value.strAttribName + " не найден в базе магазина.");
                            // нужно поставить отметки о необходимости синхронизации
                            dynamic find = connect.Справочники.Номенклатура.НайтиПоРеквизиту("КодДляСинхронизации", val.Value.iTovarCode);
                            dynamic tovar_ = find.ПолучитьОбъект();

                            tovar_.ДатаИзменения = DateTime.Now;
                            tovar_.ДатаИзмененияИзображений = DateTime.Now;
                            tovar_.ДатаИзмененияАттрибутов = DateTime.Now;

                            tovar_.Записать();

                            Marshal.FinalReleaseComObject(find);
                            Marshal.FinalReleaseComObject(tovar_);
                        }
                    }
                }

                foreach (KeyValuePair<string, PairTovarAttribInfo> valShop in dictPairTovarAttribInfoShop)
                {
                    PairTovarAttribInfo val;
                    if (!dictPairTovarAttribInfo.TryGetValue(valShop.Key, out val))
                    {
                        if (!valShop.Value.bTovarDeleted || !valShop.Value.bAttribDeleted)
                        {
                            bgVerifyTovar.ReportProgress(1, "Товар " + valShop.Value.strTovarName + ", " + valShop.Value.strAttribName + " не найден в основной базе.");

                            // нужно поставить отметки о необходимости синхронизации
                            dynamic find = connectShop.Справочники.Номенклатура.НайтиПоРеквизиту("КодДляСинхронизации", valShop.Value.iTovarCode);
                            dynamic tovar_ = find.ПолучитьОбъект();

                            tovar_.ДатаИзменения = DateTime.Now;
                            tovar_.ДатаИзмененияИзображений = DateTime.Now;
                            tovar_.ДатаИзмененияАттрибутов = DateTime.Now;

                            tovar_.Записать();

                            Marshal.FinalReleaseComObject(find);
                            Marshal.FinalReleaseComObject(tovar_);
                        }
                    }
                }

                bgVerifyTovar.ReportProgress(2, "Проверка номенклатур завершена.");

            }
            catch (Exception ex)
            {
                bgVerifyTovar.ReportProgress(1, "Ошибка при синхронизации: " + ex.Message);
                Cursor.Current = Cursors.Default;
                return;
            }

            Cursor.Current = Cursors.Default;
        }

        private void bgVerifyTovar_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            if (e.ProgressPercentage == -1)
            {
                ListViewItem vi = listView1.Items[listView1.Items.Count - 1];
                vi.Text = (string)e.UserState;
                listView1.EnsureVisible(listView1.Items.Count - 1);
            }
            else
            {
                ListViewItem vi = listView1.Items.Add((string)e.UserState);
                if ((string)e.UserState != "") vi.ImageIndex = e.ProgressPercentage;
                listView1.EnsureVisible(listView1.Items.Count - 1);
            }
        }

        private void bgVerifyTovar_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (listView1.Items.Count > 0)
                listView1.EnsureVisible(listView1.Items.Count - 1);

            gbMain.Enabled = true;
            gbShop.Enabled = true;
            gbSync.Enabled = true;
            btnRefreshDataSite.Enabled = true;
            btnYML.Enabled = true;
        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void label4_DoubleClick(object sender, EventArgs e)
        {
            if (MessageBox.Show("Проверить номенклатуры?", "Предупреждение", MessageBoxButtons.OKCancel) == System.Windows.Forms.DialogResult.OK)
            {
                gbMain.Enabled = false;
                gbShop.Enabled = false;
                gbSync.Enabled = false;
                btnRefreshDataSite.Enabled = false;
                btnYML.Enabled = false;

                listView1.Items.Clear();

                m_strSelectedShop = cbShop.Items[cbShop.SelectedIndex].ToString();

                if (ref_shop != null) Marshal.FinalReleaseComObject(ref_shop);

                ref_shop = connect.Справочники.Магазины.НайтиПоНаименованию(m_strSelectedShop, true);
                if (ref_shop == connect.Справочники.Магазины.ПустаяСсылка())
                    MessageBox.Show("Необходимо выбрать магазин.");

                bgVerifyTovar.RunWorkerAsync();
            }
        }

        private void bVerifyInvoice_DoWork(object sender, DoWorkEventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;

            try
            {
                // рассчитываем кол-во номенклатур на основной базе для выбранного магазина
                bgVerifyInvoice.ReportProgress(0, "");

                listVerifyInvoiceShop.Clear();

                {
                    bgVerifyInvoice.ReportProgress(0, "Получение данных по документам прихода.");

                    dynamic queryIn = connectShop.NewObject("Запрос");
                    queryIn.Текст = "ВЫБРАТЬ Номер, Дата, Проведен, ПометкаУдаления, КодДляСинхронизации, СуммаДокумента ИЗ Документ.ПоступлениеТоваров ";

                    dynamic objectsIn2 = queryIn.Выполнить();
                    dynamic objectsIn = objectsIn2.Выбрать();

                    VerifyInvoiceInfo vii;
                    int iCnt = 1;
                    int iCntAll = objectsIn.Количество();
                    while (objectsIn.Следующий())
                    {
                        vii = new VerifyInvoiceInfo();
                        vii.strCode = objectsIn.Номер;
                        vii.strCodeSync = objectsIn.КодДляСинхронизации;
                        vii.date = objectsIn.Дата;
                        vii.bProvodka = objectsIn.Проведен;
                        vii.bDeleted = objectsIn.ПометкаУдаления;
                        vii.fCost = objectsIn.СуммаДокумента;
                        vii.type = VerifyInvoiceInfo.InvoiceType.InvoiceIn;

                        listVerifyInvoiceShop.Add(vii);

                        bgVerifyInvoice.ReportProgress(-1, "Получение данных по документам прихода (" + iCnt.ToString() + " из " + iCntAll.ToString() + ").");
                        iCnt++;
                    }

                    Marshal.FinalReleaseComObject(objectsIn2);
                    Marshal.FinalReleaseComObject(objectsIn);
                    Marshal.FinalReleaseComObject(queryIn);
                }

                {
                    bgVerifyInvoice.ReportProgress(0, "Получение данных по документам реализации.");

                    dynamic queryIn = connectShop.NewObject("Запрос");
                    queryIn.Текст = "ВЫБРАТЬ Номер, Дата, Проведен, ПометкаУдаления, КодДляСинхронизации, СуммаДокумента ИЗ Документ.РеализацияТоваров ";

                    dynamic objectsIn2 = queryIn.Выполнить();
                    dynamic objectsIn = objectsIn2.Выбрать();

                    VerifyInvoiceInfo vii;
                    int iCnt = 1;
                    int iCntAll = objectsIn.Количество();
                    while (objectsIn.Следующий())
                    {
                        vii = new VerifyInvoiceInfo();
                        vii.strCode = objectsIn.Номер;
                        vii.strCodeSync = objectsIn.КодДляСинхронизации;
                        vii.date = objectsIn.Дата;
                        vii.bProvodka = objectsIn.Проведен;
                        vii.bDeleted = objectsIn.ПометкаУдаления;
                        vii.fCost = objectsIn.СуммаДокумента;
                        vii.type = VerifyInvoiceInfo.InvoiceType.InvoiceOut;

                        listVerifyInvoiceShop.Add(vii);

                        bgVerifyInvoice.ReportProgress(-1, "Получение данных по документам реализации (" + iCnt.ToString() + " из " + iCntAll.ToString() + ").");
                        iCnt++;
                    }

                    Marshal.FinalReleaseComObject(objectsIn2);
                    Marshal.FinalReleaseComObject(objectsIn);
                    Marshal.FinalReleaseComObject(queryIn);
                }

                {
                    bgVerifyInvoice.ReportProgress(0, "Получение данных по документам возврата.");

                    dynamic queryIn = connectShop.NewObject("Запрос");
                    queryIn.Текст = "ВЫБРАТЬ Номер, Дата, Проведен, ПометкаУдаления, КодДляСинхронизации, СуммаДокумента ИЗ Документ.ВозвратТоваровОтПокупателя ";

                    dynamic objectsIn2 = queryIn.Выполнить();
                    dynamic objectsIn = objectsIn2.Выбрать();

                    VerifyInvoiceInfo vii;
                    int iCnt = 1;
                    int iCntAll = objectsIn.Количество();
                    while (objectsIn.Следующий())
                    {
                        vii = new VerifyInvoiceInfo();
                        vii.strCode = objectsIn.Номер;
                        vii.strCodeSync = objectsIn.КодДляСинхронизации;
                        vii.date = objectsIn.Дата;
                        vii.bProvodka = objectsIn.Проведен;
                        vii.bDeleted = objectsIn.ПометкаУдаления;
                        vii.fCost = objectsIn.СуммаДокумента;
                        vii.type = VerifyInvoiceInfo.InvoiceType.Return;

                        listVerifyInvoiceShop.Add(vii);

                        bgVerifyInvoice.ReportProgress(-1, "Получение данных по документам возврата (" + iCnt.ToString() + " из " + iCntAll.ToString() + ").");
                        iCnt++;
                    }

                    Marshal.FinalReleaseComObject(objectsIn2);
                    Marshal.FinalReleaseComObject(objectsIn);
                    Marshal.FinalReleaseComObject(queryIn);
                }

                {
                    bgVerifyInvoice.ReportProgress(0, "Получение данных по документам оприходования.");

                    dynamic queryIn = connectShop.NewObject("Запрос");
                    queryIn.Текст = "ВЫБРАТЬ Номер, Дата, Проведен, ПометкаУдаления, КодДляСинхронизации, СуммаДокумента ИЗ Документ.ОприходованиеТоваров ГДЕ Фиктивный = ложь";

                    dynamic objectsIn2 = queryIn.Выполнить();
                    dynamic objectsIn = objectsIn2.Выбрать();

                    VerifyInvoiceInfo vii;
                    int iCnt = 1;
                    int iCntAll = objectsIn.Количество();
                    while (objectsIn.Следующий())
                    {
                        vii = new VerifyInvoiceInfo();
                        vii.strCode = objectsIn.Номер;
                        vii.strCodeSync = objectsIn.КодДляСинхронизации;
                        vii.date = objectsIn.Дата;
                        vii.bProvodka = objectsIn.Проведен;
                        vii.bDeleted = objectsIn.ПометкаУдаления;
                        vii.fCost = objectsIn.СуммаДокумента;
                        vii.type = VerifyInvoiceInfo.InvoiceType.MoveIn;

                        listVerifyInvoiceShop.Add(vii);

                        bgVerifyInvoice.ReportProgress(-1, "Получение данных по документам оприходования (" + iCnt.ToString() + " из " + iCntAll.ToString() + ").");
                        iCnt++;
                    }

                    Marshal.FinalReleaseComObject(objectsIn2);
                    Marshal.FinalReleaseComObject(objectsIn);
                    Marshal.FinalReleaseComObject(queryIn);
                }

                {
                    bgVerifyInvoice.ReportProgress(0, "Получение данных по документам списания.");

                    dynamic queryIn = connectShop.NewObject("Запрос");
                    queryIn.Текст = "ВЫБРАТЬ Номер, Дата, Проведен, ПометкаУдаления, КодДляСинхронизации, СуммаДокумента ИЗ Документ.СписаниеТоваров ГДЕ Фиктивный = ложь";

                    dynamic objectsIn2 = queryIn.Выполнить();
                    dynamic objectsIn = objectsIn2.Выбрать();

                    VerifyInvoiceInfo vii;
                    int iCnt = 1;
                    int iCntAll = objectsIn.Количество();
                    while (objectsIn.Следующий())
                    {
                        vii = new VerifyInvoiceInfo();
                        vii.strCode = objectsIn.Номер;
                        vii.strCodeSync = objectsIn.КодДляСинхронизации;
                        vii.date = objectsIn.Дата;
                        vii.bProvodka = objectsIn.Проведен;
                        vii.bDeleted = objectsIn.ПометкаУдаления;
                        vii.fCost = objectsIn.СуммаДокумента;
                        vii.type = VerifyInvoiceInfo.InvoiceType.MoveOut;

                        listVerifyInvoiceShop.Add(vii);

                        bgVerifyInvoice.ReportProgress(-1, "Получение данных по документам списания (" + iCnt.ToString() + " из " + iCntAll.ToString() + ").");
                        iCnt++;
                    }

                    Marshal.FinalReleaseComObject(objectsIn2);
                    Marshal.FinalReleaseComObject(objectsIn);
                    Marshal.FinalReleaseComObject(queryIn);
                }

                for (int i = 0; i < listVerifyInvoiceShop.Count; i++)
                {
                    switch (listVerifyInvoiceShop[i].type)
                    {
                        case VerifyInvoiceInfo.InvoiceType.InvoiceIn:
                            {
                                dynamic findDoc = connect.Документы.ПоступлениеТоваров.НайтиПоРеквизиту("КодДляСинхронизации", listVerifyInvoiceShop[i].strCodeSync);
                                if (findDoc == connect.Документы.ПоступлениеТоваров.ПустаяСсылка())
                                {
                                    // если новый и не проведен, тогда не переносим
                                    if (listVerifyInvoiceShop[i].bProvodka)
                                    {
                                        bgVerifyInvoice.ReportProgress(1, "Документ поступления №" + listVerifyInvoiceShop[i].strCode + " от " + listVerifyInvoiceShop[i].date.ToShortDateString() + " не найден в основной базе.");
                                    }
                                }
                                else
                                {
                                    if (findDoc.Дата != listVerifyInvoiceShop[i].date)
                                    {
                                        DateTime date = findDoc.Дата;
                                        bgVerifyInvoice.ReportProgress(1, "Документ поступления №" + listVerifyInvoiceShop[i].strCode + " от " + listVerifyInvoiceShop[i].date.ToShortDateString() + ": различается дата (" + date.ToShortDateString() + " --- " + listVerifyInvoiceShop[i].date.ToShortDateString() + ").");
                                    }
                                    if (findDoc.Проведен != listVerifyInvoiceShop[i].bProvodka)
                                    {
                                        bgVerifyInvoice.ReportProgress(1, "Документ поступления №" + listVerifyInvoiceShop[i].strCode + " от " + listVerifyInvoiceShop[i].date.ToShortDateString() + ": различается проводка.");
                                    }
                                    if (findDoc.ПометкаУдаления != listVerifyInvoiceShop[i].bDeleted)
                                    {
                                        bgVerifyInvoice.ReportProgress(1, "Документ поступления №" + listVerifyInvoiceShop[i].strCode + " от " + listVerifyInvoiceShop[i].date.ToShortDateString() + ": различается пометка удаления.");
                                    }
                                    if (findDoc.СуммаДокумента != listVerifyInvoiceShop[i].fCost)
                                    {
                                        double fCost_ = findDoc.СуммаДокумента;
                                        bgVerifyInvoice.ReportProgress(1, "Документ поступления №" + listVerifyInvoiceShop[i].strCode + " от " + listVerifyInvoiceShop[i].date.ToShortDateString() + ": различается стоимость (" + fCost_.ToString() + " --- " + listVerifyInvoiceShop[i].fCost.ToString() + ").");
                                    }
                                }

                                Marshal.FinalReleaseComObject(findDoc);
                            }
                            break;
                        case VerifyInvoiceInfo.InvoiceType.InvoiceOut:
                            {
                                dynamic findDoc = connect.Документы.РеализацияТоваров.НайтиПоРеквизиту("КодДляСинхронизации", listVerifyInvoiceShop[i].strCodeSync);
                                if (findDoc == connect.Документы.РеализацияТоваров.ПустаяСсылка())
                                {
                                    // если новый и не проведен, тогда не переносим
                                    if (listVerifyInvoiceShop[i].bProvodka)
                                    {
                                        bgVerifyInvoice.ReportProgress(1, "Документ реализации №" + listVerifyInvoiceShop[i].strCode + " от " + listVerifyInvoiceShop[i].date.ToShortDateString() + " не найден в основной базе.");
                                    }
                                }
                                else
                                {
                                    if (findDoc.Дата != listVerifyInvoiceShop[i].date)
                                    {
                                        DateTime date = findDoc.Дата;
                                        bgVerifyInvoice.ReportProgress(1, "Документ реализации №" + listVerifyInvoiceShop[i].strCode + " от " + listVerifyInvoiceShop[i].date.ToShortDateString() + ": различается дата (" + date.ToShortDateString() + " --- " + listVerifyInvoiceShop[i].date.ToShortDateString() + ").");
                                    }
                                    if (findDoc.Проведен != listVerifyInvoiceShop[i].bProvodka)
                                    {
                                        bgVerifyInvoice.ReportProgress(1, "Документ реализации №" + listVerifyInvoiceShop[i].strCode + " от " + listVerifyInvoiceShop[i].date.ToShortDateString() + ": различается проводка.");
                                    }
                                    if (findDoc.ПометкаУдаления != listVerifyInvoiceShop[i].bDeleted)
                                    {
                                        bgVerifyInvoice.ReportProgress(1, "Документ реализации №" + listVerifyInvoiceShop[i].strCode + " от " + listVerifyInvoiceShop[i].date.ToShortDateString() + ": различается пометка удаления.");
                                    }
                                    if (findDoc.СуммаДокумента != listVerifyInvoiceShop[i].fCost)
                                    {
                                        double fCost_ = findDoc.СуммаДокумента;
                                        bgVerifyInvoice.ReportProgress(1, "Документ реализации №" + listVerifyInvoiceShop[i].strCode + " от " + listVerifyInvoiceShop[i].date.ToShortDateString() + ": различается стоимость (" + fCost_.ToString() + " --- " + listVerifyInvoiceShop[i].fCost.ToString() + ").");
                                    }
                                }

                                Marshal.FinalReleaseComObject(findDoc);
                            }
                            break;
                        case VerifyInvoiceInfo.InvoiceType.Return:
                            {
                                dynamic findDoc = connect.Документы.ВозвратТоваровОтПокупателя.НайтиПоРеквизиту("КодДляСинхронизации", listVerifyInvoiceShop[i].strCodeSync);
                                if (findDoc == connect.Документы.ВозвратТоваровОтПокупателя.ПустаяСсылка())
                                {
                                    // если новый и не проведен, тогда не переносим
                                    if (listVerifyInvoiceShop[i].bProvodka)
                                    {
                                        bgVerifyInvoice.ReportProgress(1, "Документ возврата №" + listVerifyInvoiceShop[i].strCode + " от " + listVerifyInvoiceShop[i].date.ToShortDateString() + " не найден в основной базе.");
                                    }
                                }
                                else
                                {
                                    if (findDoc.Дата != listVerifyInvoiceShop[i].date)
                                    {
                                        DateTime date = findDoc.Дата;
                                        bgVerifyInvoice.ReportProgress(1, "Документ возврата №" + listVerifyInvoiceShop[i].strCode + " от " + listVerifyInvoiceShop[i].date.ToShortDateString() + ": различается дата (" + date.ToShortDateString() + " --- " + listVerifyInvoiceShop[i].date.ToShortDateString() + ").");
                                    }
                                    if (findDoc.Проведен != listVerifyInvoiceShop[i].bProvodka)
                                    {
                                        bgVerifyInvoice.ReportProgress(1, "Документ возврата №" + listVerifyInvoiceShop[i].strCode + " от " + listVerifyInvoiceShop[i].date.ToShortDateString() + ": различается проводка.");
                                    }
                                    if (findDoc.ПометкаУдаления != listVerifyInvoiceShop[i].bDeleted)
                                    {
                                        bgVerifyInvoice.ReportProgress(1, "Документ возврата №" + listVerifyInvoiceShop[i].strCode + " от " + listVerifyInvoiceShop[i].date.ToShortDateString() + ": различается пометка удаления.");
                                    }
                                    if (findDoc.СуммаДокумента != listVerifyInvoiceShop[i].fCost)
                                    {
                                        double fCost_ = findDoc.СуммаДокумента;
                                        bgVerifyInvoice.ReportProgress(1, "Документ возврата №" + listVerifyInvoiceShop[i].strCode + " от " + listVerifyInvoiceShop[i].date.ToShortDateString() + ": различается стоимость (" + fCost_.ToString() + " --- " + listVerifyInvoiceShop[i].fCost.ToString() + ").");
                                    }
                                }

                                Marshal.FinalReleaseComObject(findDoc);
                            }
                            break;
                        case VerifyInvoiceInfo.InvoiceType.MoveIn:
                            {
                                dynamic findDoc = connect.Документы.ОприходованиеТоваров.НайтиПоРеквизиту("КодДляСинхронизации", listVerifyInvoiceShop[i].strCodeSync);
                                if (findDoc == connect.Документы.ОприходованиеТоваров.ПустаяСсылка())
                                {
                                    // если новый и не проведен, тогда не переносим
                                    if (listVerifyInvoiceShop[i].bProvodka)
                                    {
                                        bgVerifyInvoice.ReportProgress(1, "Документ оприходования №" + listVerifyInvoiceShop[i].strCode + " от " + listVerifyInvoiceShop[i].date.ToShortDateString() + " не найден в основной базе.");
                                    }
                                }
                                else
                                {
                                    if (findDoc.Дата != listVerifyInvoiceShop[i].date)
                                    {
                                        DateTime date = findDoc.Дата;
                                        bgVerifyInvoice.ReportProgress(1, "Документ оприходования №" + listVerifyInvoiceShop[i].strCode + " от " + listVerifyInvoiceShop[i].date.ToShortDateString() + ": различается дата (" + date.ToShortDateString() + " --- " + listVerifyInvoiceShop[i].date.ToShortDateString() + ").");
                                    }
                                    if (findDoc.Проведен != listVerifyInvoiceShop[i].bProvodka)
                                    {
                                        bgVerifyInvoice.ReportProgress(1, "Документ оприходования №" + listVerifyInvoiceShop[i].strCode + " от " + listVerifyInvoiceShop[i].date.ToShortDateString() + ": различается проводка.");
                                    }
                                    if (findDoc.ПометкаУдаления != listVerifyInvoiceShop[i].bDeleted)
                                    {
                                        bgVerifyInvoice.ReportProgress(1, "Документ оприходования №" + listVerifyInvoiceShop[i].strCode + " от " + listVerifyInvoiceShop[i].date.ToShortDateString() + ": различается пометка удаления.");
                                    }
                                    if (findDoc.СуммаДокумента != listVerifyInvoiceShop[i].fCost)
                                    {
                                        double fCost_ = findDoc.СуммаДокумента;
                                        bgVerifyInvoice.ReportProgress(1, "Документ оприходования №" + listVerifyInvoiceShop[i].strCode + " от " + listVerifyInvoiceShop[i].date.ToShortDateString() + ": различается стоимость (" + fCost_.ToString() + " --- " + listVerifyInvoiceShop[i].fCost.ToString() + ").");
                                    }
                                }

                                Marshal.FinalReleaseComObject(findDoc);
                            }
                            break;
                        case VerifyInvoiceInfo.InvoiceType.MoveOut:
                            {
                                dynamic findDoc = connect.Документы.СписаниеТоваров.НайтиПоРеквизиту("КодДляСинхронизации", listVerifyInvoiceShop[i].strCodeSync);
                                if (findDoc == connect.Документы.СписаниеТоваров.ПустаяСсылка())
                                {
                                    // если новый и не проведен, тогда не переносим
                                    if (listVerifyInvoiceShop[i].bProvodka)
                                    {
                                        bgVerifyInvoice.ReportProgress(1, "Документ списания №" + listVerifyInvoiceShop[i].strCode + " от " + listVerifyInvoiceShop[i].date.ToShortDateString() + " не найден в основной базе.");
                                    }
                                }
                                else
                                {
                                    if (findDoc.Дата != listVerifyInvoiceShop[i].date)
                                    {
                                        DateTime date = findDoc.Дата;
                                        bgVerifyInvoice.ReportProgress(1, "Документ списания №" + listVerifyInvoiceShop[i].strCode + " от " + listVerifyInvoiceShop[i].date.ToShortDateString() + ": различается дата (" + date.ToShortDateString() + " --- " + listVerifyInvoiceShop[i].date.ToShortDateString() + ").");
                                    }
                                    if (findDoc.Проведен != listVerifyInvoiceShop[i].bProvodka)
                                    {
                                        bgVerifyInvoice.ReportProgress(1, "Документ списания №" + listVerifyInvoiceShop[i].strCode + " от " + listVerifyInvoiceShop[i].date.ToShortDateString() + ": различается проводка.");
                                    }
                                    if (findDoc.ПометкаУдаления != listVerifyInvoiceShop[i].bDeleted)
                                    {
                                        bgVerifyInvoice.ReportProgress(1, "Документ списания №" + listVerifyInvoiceShop[i].strCode + " от " + listVerifyInvoiceShop[i].date.ToShortDateString() + ": различается пометка удаления.");
                                    }
                                    if (findDoc.СуммаДокумента != listVerifyInvoiceShop[i].fCost)
                                    {
                                        double fCost_ = findDoc.СуммаДокумента;
                                        bgVerifyInvoice.ReportProgress(1, "Документ списания №" + listVerifyInvoiceShop[i].strCode + " от " + listVerifyInvoiceShop[i].date.ToShortDateString() + ": различается стоимость (" + fCost_.ToString() + " --- " + listVerifyInvoiceShop[i].fCost.ToString() + ").");
                                    }
                                }

                                Marshal.FinalReleaseComObject(findDoc);
                            }
                            break;
                    }
                }

                bgVerifyInvoice.ReportProgress(2, "Проверка накладных завершена.");

            }
            catch (Exception ex)
            {
                bgVerifyInvoice.ReportProgress(1, "Ошибка при синхронизации: " + ex.Message);
                Cursor.Current = Cursors.Default;
                return;
            }

            Cursor.Current = Cursors.Default;
        }

        private void label6_DoubleClick(object sender, EventArgs e)
        {
            if (MessageBox.Show("Проверить накладные?", "Предупреждение", MessageBoxButtons.OKCancel) == System.Windows.Forms.DialogResult.OK)
            {
                gbMain.Enabled = false;
                gbShop.Enabled = false;
                gbSync.Enabled = false;
                btnRefreshDataSite.Enabled = false;
                btnYML.Enabled = false;

                listView1.Items.Clear();

                m_strSelectedShop = cbShop.Items[cbShop.SelectedIndex].ToString();

                if (ref_shop != null) Marshal.FinalReleaseComObject(ref_shop);

                ref_shop = connect.Справочники.Магазины.НайтиПоНаименованию(m_strSelectedShop, true);
                if (ref_shop == connect.Справочники.Магазины.ПустаяСсылка())
                    MessageBox.Show("Необходимо выбрать магазин.");

                bgVerifyInvoice.RunWorkerAsync();
            }
        }

        private void backgroundWorker5_DoWork(object sender, DoWorkEventArgs e)
        {
            CreateYML();
        }

        private void CreateYML()
        {
            Cursor.Current = Cursors.WaitCursor;

            try
            {
                // рассчитываем кол-во номенклатур на основной базе для выбранного магазина
                backgroundWorker5.ReportProgress(0, "");


                backgroundWorker5.ReportProgress(0, "Создание файла для Яндекса (данная операция может занять более 5 минут).");
                dynamic mod = connect.ExchangeData;
                string strres = mod.ВыгрузитьЯндексДанные("C:/yml_true_21m.xml");
                if (strres == "")
                    backgroundWorker5.ReportProgress(2, "Создание файла для Яндекса прошло успешно.");
                else
                    backgroundWorker5.ReportProgress(1, "Ошибка при создании файла для Яндекса: " + strres);

                backgroundWorker5.ReportProgress(0, "Сохранение файла на FTP");
                strres = mod.СохранитьФайлНаФТП("C:/yml_true_21m.xml", "yml_true_21m.xml");
                if (strres == "")
                    backgroundWorker5.ReportProgress(2, "Сохранение файла на FTP прошло успешно.");
                else
                    backgroundWorker5.ReportProgress(1, "Ошибка при сохранении файла на FTP: " + strres);

                Marshal.FinalReleaseComObject(mod);
            }
            catch (Exception ex)
            {
                backgroundWorker5.ReportProgress(1, "Ошибка при синхронизации: " + ex.Message);
                Cursor.Current = Cursors.Default;
                return;
            }


            Cursor.Current = Cursors.Default;
        }

        private void backgroundWorker5_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            if (e.ProgressPercentage == -1)
            {
                ListViewItem vi = listView1.Items[listView1.Items.Count - 1];
                vi.Text = (string)e.UserState;
                listView1.EnsureVisible(listView1.Items.Count - 1);
            }
            else
            {
                ListViewItem vi = listView1.Items.Add((string)e.UserState);
                if ((string)e.UserState != "") vi.ImageIndex = e.ProgressPercentage;
                listView1.EnsureVisible(listView1.Items.Count - 1);
            }
        }

        private void backgroundWorker5_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (listView1.Items.Count > 0)
                listView1.EnsureVisible(listView1.Items.Count - 1);

            gbMain.Enabled = true;
            gbShop.Enabled = true;
            gbSync.Enabled = true;
            if (m_bWasShopSelect)
            {
                gbSync.Enabled = true;
            }
            btnRefreshDataSite.Enabled = true;
            btnYML.Enabled = true;
        }

        private void btnYML_Click(object sender, EventArgs e)
        {
            if (m_bWasShopSelect)
            {
                gbSync.Enabled = false;
            }
            gbMain.Enabled = false;
            gbShop.Enabled = false;
            btnRefreshDataSite.Enabled = false;
            btnYML.Enabled = false;

            listView1.Items.Clear();

            backgroundWorker5.RunWorkerAsync();
        }
    }
}
