<?php
/*
 * This file is part of the infotech/yii-phpword-document-generator package.
 *
 * (c) Infotech, Ltd
 *
 * For the full copyright and license information, please view the LICENSE
 * file that was distributed with this source code.
 */

namespace Infotech\PhpWordDocumentGenerator;

use Infotech\DocumentGenerator\Renderer\RendererInterface;
use PhpOffice\PhpWord\Template;

class PhpWordRenderer implements RendererInterface
{
    /**
     * Render template with data.
     *
     * @param string $templatePath
     * @param array $data
     * @return string Rendered document as binary string
     */
    public function render($templatePath, array $data)
    {
        $doc = new Template($templatePath);

        $property = new \ReflectionProperty($doc, 'tempDocumentMainPart');
        $property->setAccessible(true);
        $xml = $property->getValue($doc);

        foreach ($data as $placeholder => $value) {

            if (false !== strpos($xml, $placeholder)) {
                list($placeholder, $value) = $this->prepareBeforeSaveValue($xml, $placeholder, $value);
                $property->setValue($doc, $xml);

                foreach ($placeholder as $key => $search) {
                    $doc->setValue($search, $value[$key]);
                }

                $xml = $property->getValue($doc);
            }
        }

        return $this->getTemporaryFileContents($doc->save());
    }

    /**
     * prepareBeforeSaveValue
     *
     * @param string $documentXML
     * @param mixed  $search
     * @param mixed  $replace
     *
     * @return array
     */
    protected function prepareBeforeSaveValue(&$documentXML, $search, $replace)
    {
        $result  = [[], []];
        $replace = (array) $replace;

        foreach ((array) $search as $keyS => $valueS) {
            if (isset($replace[$keyS])) {
                if (is_array($replace[$keyS])) {
                    $dom = new \DomDocument();
                    $dom->loadXML($documentXML);
                    $list = $this->getWPNodesList($dom, $valueS);

                    foreach ($replace[$keyS] as $keyR => $valueR) {
                        $result[0][] = ($newValueS = $this->getEnclosedSearchParam($valueS . '_' . $keyR));
                        $result[1][] = $valueR;

                        foreach ($list as $node) {
                            $newNode = $node->parentNode->insertBefore($node->cloneNode(true), $node);

                            foreach ($this->getWPNodesList($dom, $valueS, false, $newNode) as $innerNode) {
                                $innerNode->nodeValue = $newValueS;
                            }
                        }
                    }

                    foreach ($this->getWPNodesList($dom, $valueS) as $node) {
                        $node->parentNode->removeChild($node);
                    }

                    $documentXML = $dom->saveXML();

                } else {
                    $result[0][] = $this->getEnclosedSearchParam($valueS);
                    $result[1][] = $replace[$keyS];
                }
            }
        }

        return $result;
    }

    /**
     * Search for parent "<w:p>" tag
     *
     * @param \DomDocument $dom
     * @param string       $search
     * @param boolean      $isParent
     * @param \DOMNode     $node
     *
     * @return \DOMNodeList
     */
    protected function getWPNodesList(\DomDocument $dom, $search, $isParent = true, $node = null)
    {
        $list = (new \DomXPath($dom))->query(
            (isset($node) ? '.' : '')
                . '//w:t[.="' . $this->getEnclosedSearchParam($search) . '"]'
                . ($isParent ? '/../..' : ''),
            $node
        );

        return $list instanceof \DOMNodeList ? $list : new \DOMNodeList();
    }

    /**
     * getEnclosedSearchParam
     *
     * @param string  $value
     * @param boolean $check
     *
     * @return string
     */
    protected function getEnclosedSearchParam($value, $check = true)
    {
        if (!$check || !preg_match('/^\$\{\w+\}$/S', $value)) {
            $value = '${' . $value . '}';
        }

        return $value;
    }

    /**
     * @param string $filePath
     * @return string
     */
    private function getTemporaryFileContents($filePath)
    {
        $contents = file_get_contents($filePath);
        unlink($filePath);
        return $contents;
    }
}
