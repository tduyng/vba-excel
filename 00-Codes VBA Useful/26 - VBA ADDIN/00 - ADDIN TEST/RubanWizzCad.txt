<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui">
<!-- Indiquez True pour masquer tous les autres onglets standards-->
<ribbon startFromScratch="false">

<tabs>
  <!-- Crée un onglet personnalisé: -->
  <!-- L'onglet va se positionner automatiquement à la fin du ruban. -->
  <!-- Utilisez insertAfterMso="NomOngletPrédéfini" pour préciser l'emplacement de l'onglet -->
  <tab id="tabWizzcad" label="WIZZCAD" visible="true">

    <!-- Crée un groupe -->  
    <group id="grWizzcad" label="WIZZCAD">

      <!-- Crée un bouton: -->
      <!--onAction="ProcLancement" définit la macro qui va être déclenchée lorsque vous allez cliquer sur le bouton -->
    
      <!--imageMso="StartAfterPrevious" définit une image de la galerie Office qui va s'afficher sur le bouton. -->
	<!--(consultez la FAQ Excel "Comment retrouver l'ID de chaque contrôle du ruban ?" pour plus de détails). -->
      <!-- Nota: il est aussi possible d'ajouter des images externes pour personnaliser les boutons -->
      <button id="btLogin" label="Connexion" screentip="Connexion WizzCAD."
      onAction="btAiLogin" 
      supertip="Utiliser ce bouton pour la connexion avec le web WizzCAD." 
      size="large" image="WIZZCAD" />

      <!-- Crée un deuxième bouton -->	
      <button id="btImport" label="Import" screentip="Import WizzCAD."
      onAction="btAiImport" size="large" 
      supertip="Importer des données via le web WizzCAD." 
      image="IMPORT" />

	        <!-- Crée un troisième bouton -->	
      <button id="btExport" label="Export" screentip="Export WizzCAD."
      onAction="btAiExport" size="large" 
      supertip="Exporter des données via le web WizzCAD." 
      image="EXPORT" />
	  
    </group>
  </tab>

</tabs>
</ribbon>
</customUI>