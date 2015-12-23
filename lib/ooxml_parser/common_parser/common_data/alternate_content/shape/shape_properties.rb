require_relative 'shape_properties/transform_effect'
require_relative 'text_body/ooxml_text_box'
require_relative 'text_body/text_body'
module OoxmlParser
  class PresentationShapeProperties < OOXMLDocumentObject
    attr_accessor :transform, :preset, :custom, :fill, :line

    def initialize(transform = TransformEffect.new, preset = nil)
      @transform = transform
      @preset = preset
    end
  end
end
