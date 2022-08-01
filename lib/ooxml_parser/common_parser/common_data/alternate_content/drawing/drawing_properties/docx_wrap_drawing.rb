# frozen_string_literal: true

module OoxmlParser
  # Docx Wrap Drawing
  class DocxWrapDrawing < OOXMLDocumentObject
    attr_accessor :wrap_text, :distance_from_text
    # @return [Boolean] Specifies whether this floating DrawingML
    #   object is displayed behind the text of the
    #   document when the document is displayed.
    #   When a DrawingML object is displayed
    #   within a WordprocessingML document,
    #   that object can intersect with text in the
    #   document. This attribute shall determine
    #   whether the text or the object is rendered on
    #   top in case of overlapping.
    attr_reader :behind_doc

    def initialize(parent: nil)
      @wrap_text = :none
      super
    end

    # Parse DocxWrapDrawing object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [DocxWrapDrawing] result of parsing
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'behindDoc'
          @behind_doc = attribute_enabled?(value)
        end
      end

      @wrap_text = :behind if @behind_doc == true
      @wrap_text = :infront if @behind_doc == false

      node.xpath('*').each do |node_child|
        case node_child.name
        when 'wrapSquare'
          @wrap_text = :square
          @distance_from_text = DocxDrawingDistanceFromText.new(parent: self).parse(node_child)
          break
        when 'wrapTight'
          @wrap_text = :tight
          @distance_from_text = DocxDrawingDistanceFromText.new(parent: self).parse(node_child)
          break
        when 'wrapThrough'
          @wrap_text = :through
          @distance_from_text = DocxDrawingDistanceFromText.new(parent: self).parse(node_child)
          break
        when 'wrapTopAndBottom'
          @wrap_text = :topbottom
          @distance_from_text = DocxDrawingDistanceFromText.new(parent: self).parse(node_child)
          break
        end
      end
      self
    end
  end
end
