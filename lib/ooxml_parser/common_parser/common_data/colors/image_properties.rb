module OoxmlParser
  # Class for parsing ImageProperties
  class ImageProperties < OOXMLDocumentObject
    attr_accessor :alpha_modulate_fixed_effect

    # Parse ImageProperties object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [ImageProperties] result of parsing
    def parse(node)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'alphaModFix'
          @alpha_modulate_fixed_effect = (node_child.attribute('amt').value.to_f / 1_000.0).round(1)
        end
      end
      self
    end
  end
end
