<?php
/**
 * This file is part of PHPPresentation - A pure PHP library for reading and writing
 * presentations documents.
 *
 * PHPPresentation is free software distributed under the terms of the GNU Lesser
 * General Public License version 3 as published by the Free Software Foundation.
 *
 * For the full copyright and license information, please read the LICENSE
 * file that was distributed with this source code. For the full list of
 * contributors, visit https://github.com/PHPOffice/PHPPresentation/contributors.
 *
 * @see        https://github.com/PHPOffice/PHPPresentation
 *
 * @license     http://www.gnu.org/licenses/lgpl.txt LGPL version 3
 */

declare(strict_types=1);

namespace PhpOffice\PhpPresentation\Shape\Drawing;

/**
 * Drawing that keeps an SVG in the inherited data property and a PNG fallback
 * in $pngData. relationId is the PNG relationship used by a:blip. svgRelationId
 * is the SVG relationship used by asvg:svgBlip.
 */
class DualFormatBase64 extends Base64
{
    protected string $pngData = '';

    protected string $pngUniqueName;

    public string $svgRelationId = '';

    public function __construct()
    {
        parent::__construct();
        $this->pngUniqueName = md5(mt_rand(0, 9999) . time() . mt_rand(0, 9999));
    }

    public function setPngData(string $data): self
    {
        $this->pngData = $data;

        return $this;
    }

    public function getPngData(): string
    {
        return $this->pngData;
    }

    public function getPngContents(): string
    {
        [, $imageContents] = explode(';', $this->getPngData());
        [, $imageContents] = explode(',', $imageContents);

        return base64_decode($imageContents);
    }

    public function getPngIndexedFilename(): string
    {
        return $this->pngUniqueName . $this->getImageIndex() . '.png';
    }

    public function hasPngFallback(): bool
    {
        return $this->pngData !== '';
    }
}
