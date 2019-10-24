<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui">
<ribbon startFromScratch="false">
    <tabs>
      <tab id="Tab1" label="Phan m?m QLCL">
        <group id="Group1" label="Công trình">
          <menu id="Menu1" label="H? so" imageMso="OpenStartPage" size="large" itemSize="normal">
            <button id="BT1" onAction="Taomoicongtrinh" label="T?o công trình m?i" imageMso="FileNew"/>
            <button id="BT2" onAction="Mocongtrinh" label="M? công trình" imageMso="FileOpen"/>
            <button id="BT3" onAction="Luucongtrinh" label="Luu công trình" imageMso="FileSave"/>
            <button id="BT4" onAction="Caidatchung" label="Cài d?t chung" imageMso="AddInManager"/>
          </menu>
          <menu id="Menu2" label="Thông tin chung" imageMso="AccessFormWizard" size="large" itemSize="normal">
            <button id="BT5" onAction="Thongtincongtrinh" label="Thông tin công trình" imageMso="QueryShowTable"/>
            <button id="BT6" onAction="Dulieudauvao" label="D? li?u d?u vào" imageMso="FormulaMoreFunctionsMenu"/>
          </menu>
          <menu id="Menu3" label="B?ng mã áp d?ng" imageMso="ImportAccess" size="large" itemSize="normal">
            <button id="BT7" onAction="Mahieucongviec" label="Mã hi?u công vi?c" imageMso="QueryShowTable"/>
            <button id="BT8" onAction="Mahieuthinghiem" label="Mã hi?u thí nghi?m" imageMso="FormulaMoreFunctionsMenu"/>
            <button id="BT9" onAction="Tieuchuannghiemthu" label="Tiêu chu?n nghi?m thu" imageMso="QueryShowTable"/>
            <button id="BT10" onAction="CCThinghiem" label="Can c? thí nghi?m" imageMso="FormulaMoreFunctionsMenu"/>
            <button id="BT11" onAction="Phuongphapkiemtra" label="Phuong pháp ki?m tra" imageMso="QueryShowTable"/>
          </menu>
        </group>
        <group id="Group2" label="Nh?t trình thi công">
            <button id="BT12" onAction="Nhattrinh" label="Nh?t trình" imageMso="FileBackupDatabase" size="large"/>
            <button id="BT13" onAction="Thoitiet" label="Th?i ti?t, ngày ngh?" imageMso="PictureBrightnessGallery" size="large"/>
            <button id="BT14" onAction="Kiemtrangaythicong" label="Ki?m tra ngày thi công" imageMso="FilePrepareMenu" size="normal"/>
            <button id="BT15" onAction="Xuatdulieubienban" label="Xu?t d? li?u biên b?n" imageMso="DataValidation" size="normal"/>
            <button id="BT16" onAction="Xuatdulieunhatky" label="Xu?t d? li?u nh?t ký" imageMso="ReviewTrackChanges" size="normal"/>
        </group>
        <group id="Group3" label="List Biên b?n">
            <button id="BT17" onAction="ListTNDV" label="Thí nghi?m d?u vào" imageMso="OrganizationChartSelectLevel" size="normal"/>
            <button id="BT18" onAction="ListNTCV" label="Nghi?m thu công vi?c" imageMso="CreateFormInDesignView" size="normal"/>
            <button id="BT19" onAction="ListTNHT" label="Thí nghi?m hi?n tru?ng" imageMso="ModuleInsert" size="normal"/>
            <button id="BT20" onAction="ListNTGD" label="Nghi?m thu giai do?n" imageMso="FunctionsFinancialInsertGallery" size="normal"/>
          <menu id="Menu4" label="Các biên b?n khác" imageMso="FileSaveToDocumentManagementServer" size="normal" itemSize="normal">
            <button id="BT21" onAction="ListTDBT" label="Theo dõi d? bê tông" imageMso="QueryShowTable"/>
            <button id="BT22" onAction="ListKTVK" label="Ki?m tra ván khuôn" imageMso="FormulaMoreFunctionsMenu"/>
            <button id="BT23" onAction="ListKTCD" label="Ki?m tra cao d?" imageMso="QueryShowTable"/>
            <button id="BT24" onAction="ListKTK" label="Ki?m tra khác" imageMso="FormulaMoreFunctionsMenu"/>
          </menu>
        </group>
        <group id="Group4" label="Xu?t b?n Biên b?n">
            <button id="BT25" onAction="TNDV" label="Thí nghi?m d?u vào" imageMso="OrganizationChartSelectLevel" size="normal"/>
            <button id="BT26" onAction="NTCV" label="Nghi?m thu công vi?c" imageMso="CreateFormInDesignView" size="normal"/>
            <button id="BT27" onAction="TNHT" label="Thí nghi?m hi?n tru?ng" imageMso="ModuleInsert" size="normal"/>
            <button id="BT28" onAction="NTGD" label="Nghi?m thu giai do?n" imageMso="FunctionsFinancialInsertGallery" size="normal"/>
          <menu id="Menu5" label="Cac biên b?n khác" imageMso="FileSaveToDocumentManagementServer" size="normal" itemSize="normal">
            <button id="BT29" onAction="TDBT" label="Theo dõi d? bê tông" imageMso="QueryShowTable"/>
            <button id="BT30" onAction="KTVK" label="Ki?m tra ván khuôn" imageMso="FormulaMoreFunctionsMenu"/>
            <button id="BT31" onAction="KTCD" label="Ki?m tra cao d?" imageMso="QueryShowTable"/>
            <button id="BT32" onAction="KTK" label="Ki?m tra khác" imageMso="FormulaMoreFunctionsMenu"/>
          </menu>
        </group>
        <group id="Group5" label="Nh?t ký thi công">
            <button id="BT33" onAction="Caidatnhatky" label="Cài d?t nh?t ký" imageMso="ViewDocumentActionsPane" size="large"/>
            <button id="BT34" onAction="Xuatnhatky" label="Ghi nh?t ký" imageMso="FunctionsDateTimeInsertGallery" size="large"/>
        </group>
        <group id="Group6" label="Ti?n ích khác">
          <menu id="Menu6" label="Hu?ng d?n - Tác gi?" imageMso="FunctionsLogicalInsertGallery" size="large" itemSize="normal">
            <button id="BT35" onAction="Huongdansudung" label="Hu?ng d?n s? d?ng" imageMso="FunctionsLookupReferenceInsertGallery"/>
            <button id="BT36" onAction="Videohuongdan" label="Video hu?ng d?n s? d?ng" imageMso="ViewSlideShowView"/>
            <button id="BT37" onAction="About" label="About" imageMso="ContactPictureMenu"/>
          </menu>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>