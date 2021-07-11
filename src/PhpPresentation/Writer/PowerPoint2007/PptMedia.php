<?php

namespace PhpOffice\PhpPresentation\Writer\PowerPoint2007;

use PhpOffice\PhpPresentation\Shape\Drawing\AbstractDrawingAdapter;
use PhpOffice\PhpPresentation\Shape\Drawing\Gd;
use SVG\SVG;

class PptMedia extends AbstractDecoratorWriter
{
    /**
     * @return \PhpOffice\Common\Adapter\Zip\ZipInterface
     *
     * @throws \Exception
     */
    public function render()
    {
        for ($i = 0; $i < $this->getDrawingHashTable()->count(); ++$i) {
            $shape = $this->getDrawingHashTable()->getByIndex($i);
            if (!$shape instanceof AbstractDrawingAdapter) {
                continue;
            }
            $this->getZip()->addFromString('ppt/media/' . $shape->getIndexedFilename(), $shape->getContents());

            //Make a png version of the svg
            if($shape->getExtension() == 'svg') {
                //Make a PNG file for it
                $pngFileName = str_replace(' ', '_', $shape->getIndexedFilename());
                //Replace the .svg with .png
                $pngFileName = str_replace(".svg", ".png", $pngFileName);

                $image = SVG::fromString($shape->getContents());

                $oShape = new Gd();
                $oShape->setImageResource($image->toRasterImage( 200, 200 , "#FFFFFF"))
                    ->setRenderingFunction(Gd::RENDERING_PNG)
                    ->setMimeType(Gd::MIMETYPE_DEFAULT);

                $this->getZip()->addFromString('ppt/media/' . $pngFileName, $oShape->getContents());
            }
        }

        return $this->getZip();
    }
}
