<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui">
<!-- Indiquez True pour masquer tous les autres onglets standards-->
<ribbon startFromScratch="false">

<tabs>
  <!-- Crée un onglet personnalisé: -->
  <!-- L'onglet va se positionner automatiquement à la fin du ruban. -->
  <!-- Utilisez insertAfterMso="NomOngletPrédéfini" pour préciser l'emplacement de l'onglet -->
  <tab id="OngletPerso" label="OngletPerso" visible="true">

    <!-- Crée un groupe -->  
    <group id="Essai" label="Essai CustomUI">

      <!-- Crée un bouton: -->
      <!--onAction="ProcLancement" définit la macro qui va être déclenchée lorsque vous allez cliquer sur le bouton -->
    
      <!--imageMso="StartAfterPrevious" définit une image de la galerie Office qui va s'afficher sur le bouton. -->
	<!--(consultez la FAQ Excel "Comment retrouver l'ID de chaque contrôle du ruban ?" pour plus de détails). -->
      <!-- Nota: il est aussi possible d'ajouter des images externes pour personnaliser les boutons -->
      <button id="btLance01" label="Lancement" screentip="Déclenche la procédure."
      onAction="ProcLancement" 
      supertip="Utilisez ce bouton pour Lancer la macro." 
      size="large" imageMso="StartAfterPrevious" />

      <!-- Crée un deuxième bouton -->	
      <button id="btAide01" label="Aide" screentip="Consultez l'aide."
      onAction="OuvertureAide" size="large" 
      supertip="Consultez les meilleurs cours et tutoriels Office." 
      imageMso="FunctionsLogicalInsertGallery" 
      tag="http://office.developpez.com/" />

    </group>
  </tab>

</tabs>
</ribbon>
</customUI>