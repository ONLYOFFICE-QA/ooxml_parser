# frozen_string_literal: true

module OoxmlParser
  # Class for parsing Chart Style entry
  class ChartStyleEntry < OOXMLDocumentObject
    # @return [OOXMLShapeBodyProperties] body properties
    attr_reader :body_properties
    # @return [RunProperties] default run properties
    attr_reader :default_run_properties
    # @return [FillReference] effect reference
    attr_reader :effect_reference
    # @return [FillReference] fill reference
    attr_reader :fill_reference
    # @return [FontReference] font reference
    attr_reader :font_reference
    # @return [StyleMatrixReference] line style reference
    attr_reader :line_style_reference
    # @return [StyleMatrixReference] shape properties
    attr_reader :shape_properties

    # Parse Chart style entry
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [ChartStyleEntry] result of parsing
    def parse(node)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'bodyPr'
          @body_properties = OOXMLShapeBodyProperties.new(parent: self).parse(node_child)
        when 'defRPr'
          @default_run_properties = RunProperties.new(parent: self).parse(node_child)
        when 'effectRef'
          @effect_reference = StyleMatrixReference.new(parent: self).parse(node_child)
        when 'fillRef'
          @fill_reference = StyleMatrixReference.new(parent: self).parse(node_child)
        when 'fontRef'
          @font_reference = FontReference.new(parent: self).parse(node_child)
        when 'lnRef'
          @line_style_reference = StyleMatrixReference.new(parent: self).parse(node_child)
        when 'spPr'
          @shape_properties = DocxShapeProperties.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end
