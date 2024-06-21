import Bulletin from "@/components/Bulletin";
import BulletinSteps from "@/components/BulletinSteps";
import Excel from "@/components/Excel";
import { Icons } from "@/components/Icons";
import MaxWidthWrapper from "@/components/MaxWidthWrapper";
import Reviews from "@/components/Reviews";
import { buttonVariants } from "@/components/ui/button";
import { ArrowRight, Check } from "lucide-react";
import Image from "next/image";
import Link from "next/link";

const page = () => {
  return (
    <div className="bg-slate-50">
      <section>
        <MaxWidthWrapper className="pb-24 pt-10 lg:grid lg:grid-cols-3 sm:pb-32 lg:gap-x-0 xl:gap-x-8 lg:pt-24 xl:pt-32 lg:pb-52">
          <div className="col-span-2 px-6 lg:px-0 lg:pt-4">
            <div className="relative mx-auto text-center lg:text-left flex flex-col items-center lg:items-start">
              {/* <div className="absolute w-28 left-0 -top-20 hidden lg:block">
                <Image src="/logo.png" alt="logo" width={200} height={200} />
              </div> */}
              <h1 className="relative w-fit tracking-tight text-balance mt-16 font-bold !leading-tight text-gray-900 text-5xl md:text-6xl lg:text-7xl">
                Générateur de <span className="bg-primary-50 px-2 text-white">Bulletins</span>{" "}
                Scolaire Automatisé
              </h1>
              <p className="mt-8 text-lg lg:pr-10 max-w-prose text-center lg:text-left text-balance md:text-wrap">
                Simplifiez la gestion et la distribution des bulletins scolaires semestriels et
                annuels avec notre
                <span className="font-semibold">
                  {" "}
                  application innovante intégrée à Yparéo.{" "}
                </span>{" "}
                Grâce à cette solution, vous pouvez facilement importer vos documents Excel et Word
                pour générer des bulletins détaillés et personnalisés en quelques clics.
              </p>
              <ul className="mt-8 space-y-2 text-left font-medium flex flex-col items-center sm:items-start">
                <div className="space-y-2">
                  <li className="flex gap-1.5 items-center text-left">
                    <Check className="h-5 w-5 shrink-0 text-primary-50" />
                    Gain de temps
                  </li>
                  <li className="flex gap-1.5 items-center text-left">
                    <Check className="h-5 w-5 shrink-0 text-primary-50" />
                    Réduit les erreurs manuelles
                  </li>
                  <li className="flex gap-1.5 items-center text-left">
                    <Check className="h-5 w-5 shrink-0 text-primary-50" />
                    Accessibilité et facilité d&apos;utilisation
                  </li>
                </div>
              </ul>
            </div>
          </div>

          <div className="col-span-full lg:col-span-1 w-full flex justify-center px-8 sm:px-16 md:px-0 mt-32 lg:mx-0 lg:mt-20 h-fit">
            <div className="relative md:max-w-xl">
              <Excel className="w-96" imgSrc="/excel.png" />
            </div>
          </div>
        </MaxWidthWrapper>
      </section>

      {/* value proposition section*/}
      <section className="bg-slate-100 py-24">
        <MaxWidthWrapper className="flex flex-col items-center gap-4 sm:gap-32">
          <div className="flex flex-col lg:flex-row items-center gap-4 sm:gap-6">
            <h2 className="order-1 mt-2 tracking-tight text-center text-balance !leading-tight font-bold text-5xl md:text-6xl text-gray-900">
              Une utilisation{" "}
              <span className="relative px-2">
                rapide{" "}
                <Icons.underline className="hidden sm:block pointer-events-none absolute inset-x-0 -bottom-6 text-primary-50" />
              </span>{" "}
              et efficace
            </h2>
          </div>

          <BulletinSteps />
          {/* <div className="mx-auto grid max-w-2xl grid-cols-1 px-4 lg:mx-0 lg:max-w-none lg:grid-cols-2 gap-y-16">
            <div className="flex flex-auto flex-col gap-4 lg:pr-8 xl:pr-20">
              <div className="flex gap-0.5 mb-2">
                <span className="font-semibold">Téléchargement des Documents</span>
              </div>
              <div className="text-lg leading-8">
                <p>
                  Importez vos fichiers Excel et Word contenant les informations nécessaires pour
                  les bulletins scolaires. Notre système prend en charge les formats les plus
                  courants pour assurer une compatibilité maximale.
                </p>
              </div>
            </div>
            <div className="flex flex-auto flex-col gap-4 lg:pr-8 xl:pr-20">
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
            </div>
            <div className="flex flex-auto flex-col gap-4 lg:pr-8 xl:pr-20">
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
            </div>
            <div className="flex flex-auto flex-col gap-4 lg:pr-8 xl:pr-20">
              <div className="flex gap-0.5 mb-2">
                <span className="font-semibold">Accessibilité et Facilité d&apos;Utilisation</span>
              </div>
              <div className="text-lg leading-8">
                <p>
                  Notre interface intuitive permet une navigation fluide et une utilisation
                  simplifiée, même pour les utilisateurs novices. Vous pouvez gérer vos bulletins en
                  quelques étapes simples, sans besoin de compétences techniques particulières.
                </p>
              </div>
            </div>
          </div> */}
        </MaxWidthWrapper>

        <div className="pt-16">
          <Reviews />
        </div>
      </section>

      <section>
        <MaxWidthWrapper className="py-24">
          <div className="mb-12 px-6 lg:px-8">
            <div className="mx-auto max-w-2xl sm:text-center">
              <h2 className="order-1 mt-2 tracking-tight text-center text-balance !leading-tight font-bold text-5xl md:text-6xl text-gray-900">
                Télécharger vos bulletins et
                <span className="relative px-2 bg-primary-50 text-white"> obtenez les</span>
                maintenant
              </h2>
            </div>
          </div>

          <div className="mx-auto max-w-6xl px-6 lg:px-8">
            <div className="relative flex flex-col items-center md:grid grid-cols-2 gap-40">
              <div className="absolute top-[25rem] md:top-1/2 -translate-y-1/2 z-10 left-1/2 -translate-x-1/2 rotate-90 md:rotate-0">
                <Image
                  src="/arrow.png"
                  alt="Arrow"
                  width={90} // Largeur de l'image en pixels
                  height={90} // Hauteur de l'image en pixels
                />
              </div>

              <div className="relative h-80 md:h-full w-full md:justify-self-end max-w-sm rounded-xl bg-gray-900/5 ring-inset ring-gray-900/10 lg:rounded-2xl">
                <Excel
                  className="rounded-md object-cover bg-white shadow-2xl ring-1 ring-gray-900/10 h-full w-full"
                  imgSrc="/darkexcel.png"
                />
              </div>

              <Bulletin className="w-60" imgSrc="/bulletin.png" />
            </div>
          </div>

          <ul className="mx-auto mt-12 max-w-prose sm:text-lg space-y-2 w-fit">
            <div className="flex justify-center">
              <Link
                className={buttonVariants({
                  size: "lg",
                  className: "mx-auto mt-8 bg-primary-50",
                })}
                href="/configure/upload"
              >
                Générer vos bulletins <ArrowRight className="h-4 w-4 ml-1.5" />
              </Link>
            </div>
          </ul>
        </MaxWidthWrapper>
      </section>
    </div>
  );
};

export default page;
