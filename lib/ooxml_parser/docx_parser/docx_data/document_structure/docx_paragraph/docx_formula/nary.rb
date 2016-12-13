require_relative 'nary/nary_properties'
module OoxmlParser
  # Class for parsing `m:nary` object
  class Nary < OOXMLDocumentObject
    # @return [DocxFormula] top value
    attr_accessor :top_value

    # @return [DocxFormula] bottom_value
    attr_accessor :bottom_value

    # @return [NaryProperties] properties of nary
    attr_accessor :properties

    # Parse Nary
    # @param [Nokogiri::XML:Node] node with Nary
    # @return [Nary] result of parsing
    def parse(node)
      node.xpath('*').each do |nary_child_node|
        case nary_child_node.name
        when 'sub'
          @bottom_value = DocxFormula.new(parent: self).parse(nary_child_node)
        when 'sup'
          @top_value = DocxFormula.new(parent: self).parse(nary_child_node)
        when 'naryPr'
          @properties = NaryProperties.new(parent: self).parse(nary_child_node)
        end
      end
      self
    end
  end
end
