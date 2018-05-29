<?php
/**
 * @link
 * @copyright Copyright (c) 2018 mlcsthor
 * @license [New BSD License](http://www.opensource.org/licenses/bsd-license.php)
 * @author Maxime Lucas <mlcsthor@gmail.com>
 * @since 1.0
 */

namespace mlcsthor\presentation;

use PhpOffice\PhpPresentation\AbstractShape;
use PhpOffice\PhpPresentation\DocumentLayout;
use PhpOffice\PhpPresentation\DocumentProperties;
use PhpOffice\PhpPresentation\IOFactory;
use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\Style\Font;
use Yii;
use yii\base\Component;
use yii\di\Instance;
use yii\helpers\FileHelper;
use yii\i18n\Formatter;

class Presentation extends Component {
    /**
     * @var array data used to configure both the presentation and the slide
     */
    public $data;

    /**
     * @var string the HTML display when the content of a slide is empty
     */
    public $emptySlide = '';

    /**
     * @var string writer type (format type)
     *
     * Supported values:
     * - 'PowerPoint2007' (default)
     * - 'ODPresentation'
     * - 'Serialized'
     *
     * @see IOFactory
     */
    public $writerType;

    /**
     * @var bool whether presentation has been already rendered or not
     */
    protected $isRendered = false;

    /**
     * @var PhpPresentation|null presentation document representation interface
     */
    private $_document;

    /**
     * @var array|Formatter the formatter used to format model attribute values into displayable texts.
     * This can be either an instance of [[Formatter]] or an configuration array for creating the [[Formatter]]
     * instance. If this property is not set, the "formatter" application component will be used.
     */
    private $_formatter;

    /**
     * @return PhpPresentation presentation document representation interface
     */
    public function getDocument() {
        if (!is_object($this->_document)) {
            $this->_document = new PhpPresentation();
        }

        return $this->_document;
    }

    /**
     * @param PhpPresentation|null $document presentation document representation interface
     */
    public function setDocument($document) {
        $this->_document = $document;
    }

    /**
     * @throws \yii\base\InvalidConfigException
     * @return Formatter formatter instance
     */
    public function getFormatter() {
        if (!is_object($this->_formatter)) {
            if ($this->_formatter === null) {
                $this->_formatter = Yii::$app->getFormatter();
            } else {
                $this->_formatter = Instance::ensure($this->_formatter, Formatter::class);
            }
        }

        return $this->_formatter;
    }

    /**
     * @param array|Formatter $formatter formatter instance
     */
    public function setFormatter($formatter) {
        $this->_formatter = $formatter;
    }

    /**
     * Sets presentation document properties
     * 
     * @param array $properties list of document properties in format: name => value
     * @return $this self reference
     * @see DocumentProperties
     */
    public function setDocumentProperties($properties) {
        $documentProperties = $this->getDocument()->getDocumentProperties();

        foreach ($properties as $name => $value) {
            $method = 'set' . ucfirst($name);
            call_user_func([$documentProperties, $method], $value);
        }

        return $this;
    }

    /**
     * Sets given font properties
     *
     * @param Font $font font to configure
     * @param array $properties list of font properties in format : name => value
     * @return $this self reference
     * @see Font
     */
    public function setFontProperties($font, $properties) {
        foreach ($properties as $name => $value) {
            $method = 'set' . ucfirst($name);
            call_user_func([$font, $method], $value);
        }

        return $this;
    }

    /**
     * Sets given shape properties
     *
     * @param AbstractShape $shape shape to configure
     * @param array $properties list of shape properties in format : name => value
     * @return $this self reference
     */
    public function setShapeProperties($shape, $properties) {
        foreach ($properties as $name => $value) {
            $method = 'set' . ucfirst($name);
            call_user_func([$shape, $method], $value);
        }

        return $this;
    }

    /**
     * Configures (re-configures) this presentation with the property values
     *
     * @param array $properties the property initial values given in terms of name-value pairs
     * @return $this self reference
     */
    public function configure($properties) {
        Yii::configure($this, $properties);

        return $this;
    }

    /**
     * Performs actual document composition
     *
     * @return $this self reference
     * @throws \Exception
     */
    public function render() {
        $document = $this->getDocument();

        $this->setDocumentProperties($this->data['document'] ?? []);

        if (isset($this->data['layout'])) {
            if (is_array($this->data['layout'])) {
                $this->getDocument()->getLayout()->setCX($this->data['layout']['x'], $this->data['units'] ?? DocumentLayout::UNIT_EMU);
                $this->getDocument()->getLayout()->setCX($this->data['layout']['y'], $this->data['units'] ?? DocumentLayout::UNIT_EMU);
            } else {
                $this->getDocument()->getLayout()->setDocumentLayout($this->data['layout']);
            }
        }

        foreach ($this->data['slides'] as $slideData) {
            $slide = $document->getActiveSlide();

            $slide->setName($slideData['name'] ?? null);

            foreach ($slideData['content'] as $shapeData) {
                $shape = $slide->createRichTextShape();

                $textData = $shapeData['text'];
                unset($shapeData['text']);

                if (is_array($textData)) {
                    $text = $shape->createText($textData['content']);
                    $this->setFontProperties($text->getFont(), $textData['font'] ?? []);
                } else {
                    $shape->createText($textData);
                }

                $this->setShapeProperties($shape, $shapeData);
            }

            $document->createSlide();
            $document->setActiveSlideIndex($document->getActiveSlideIndex() + 1);
        }

        $document->removeSlideByIndex($document->getActiveSlideIndex());
        $this->isRendered = true;

        return $this;
    }

    /**
     * @param string $filename name of the output file
     * @throws \yii\base\Exception
     * @throws \Exception
     */
    public function save($filename) {
        if (!$this->isRendered) {
            $this->render();
        }

        $filename = Yii::getAlias($filename);

        $writerType = $this->writerType;

        if ($writerType === null) {
            $writerType = 'PowerPoint2007';
        }

        $fileDir = pathinfo($filename, PATHINFO_DIRNAME);
        FileHelper::createDirectory($fileDir);

        $writer = IOFactory::createWriter($this->getDocument(), $writerType);
        $writer->save($filename);
    }

    /**
     * Sends the rendered content as a file to the browser
     *
     * @param string $attachmentName the file name shown to the user
     * @param array $options additional options for sending the file
     * @return \yii\web\Response the response object
     * @throws \yii\web\RangeNotSatisfiableHttpException
     * @throws \Exception
     */
    public function send($attachmentName, $options = []) {
        if (!$this->isRendered) {
            $this->render();
        }

        $writerType = $this->writerType;

        if ($writerType === null) {
            $fileExtension = strtolower(pathinfo($attachmentName, PATHINFO_EXTENSION));
            $writerType = ucfirst($fileExtension);
        }

        $tmpResource = tmpfile();

        if ($tmpResource === false) {
            throw new \RuntimeException('Unable to create temporary file.');
        }

        $tmpResourceMetaData = stream_get_meta_data($tmpResource);
        $tmpFileName = $tmpResourceMetaData['uri'];

        $writer = IOFactory::createWriter($this->getDocument(), $writerType);
        $writer->save($tmpFileName);
        unset($writer);

        return Yii::$app->getResponse()->sendStreamAsFile($tmpResource, $attachmentName, $options);
    }
}