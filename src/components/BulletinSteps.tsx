import {
  Carousel,
  CarouselContent,
  CarouselItem,
  CarouselNext,
  CarouselPrevious,
} from "./ui/carousel";

const BulletinSteps = () => {
  return (
    <div className="mx-auto max-w-2xl px-4">
      <div className="flex flex-auto flex-col gap-4">
        <Carousel>
          <CarouselContent>
            <CarouselItem className="text-justify">
              <div className="flex gap-0.5 mb-4">
                <span className="font-semibold">Téléchargement des Documents</span>
              </div>
              <div className="text-lg leading-8">
                <p>
                  Importez vos fichiers Excel et Word contenant les informations nécessaires pour
                  les bulletins scolaires. Notre système prend en charge les formats les plus
                  courants pour assurer une compatibilité maximale.
                </p>
              </div>
            </CarouselItem>
            <CarouselItem className="text-justify">
              <div className="flex gap-0.5 mb-2">
                <span className="font-semibold">Génération Automatique des Bulletins</span>
              </div>
              <div className="text-lg leading-8">
                <p>
                  Une fois vos fichiers importés, notre application traite les données pour générer
                  des bulletins détaillés pour chaque élève. Les bulletins sont personnalisés selon
                  les informations fournies et peuvent être visualisés avant validation.
                </p>
              </div>
            </CarouselItem>
            <CarouselItem className="text-justify">
              {" "}
              <div className="flex gap-0.5 mb-2">
                <span className="font-semibold">Téléchargement et Archivage</span>
              </div>
              <div className="text-lg leading-8">
                <p>
                  Après validation, vous pouvez télécharger les bulletins générés dans un fichier
                  .zip pour une gestion facile et rapide. En outre, les bulletins sont
                  automatiquement envoyés dans les dossiers des apprenants sur Yparéo, assurant
                  ainsi une organisation optimale et une distribution sans effort.
                </p>
              </div>
            </CarouselItem>
          </CarouselContent>
          <CarouselPrevious />
          <CarouselNext />
        </Carousel>
      </div>
    </div>
  );
};

export default BulletinSteps;
