require_relative 'nary/nary_properties'
module OoxmlParser
  # Class for parsing `m:nary` object
  class Nary
    # @return [DocxFormula] top value
    attr_accessor :top_value

    # @return [DocxFormula] bottom_value
    attr_accessor :bottom_value

    # @return [NaryProperties] properties of nary
    attr_accessor :properties

    # Parse Nary
    # @param [Nokogiri::XML:Node] node with Nary
    # @return [Nary] result of parsing
    def self.parse(node)
      nary = Nary.new
      node.xpath('*').each do |nary_child_node|
        case nary_child_node.name
        when 'sub'
          nary.bottom_value = DocxFormula.parse(nary_child_node)
        when 'sup'
          nary.top_value = DocxFormula.parse(nary_child_node)
        when 'naryPr'
          nary.properties = NaryProperties.parse(nary_child_node)
        end
      end
      nary
    end
  end
end
