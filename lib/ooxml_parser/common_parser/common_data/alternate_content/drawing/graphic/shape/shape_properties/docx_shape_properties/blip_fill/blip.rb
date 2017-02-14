require_relative 'blip/alpha_mod_fix'

module OoxmlParser
  # Class for parsing `a:blip` tag
  class Blip < OOXMLDocumentObject
    attr_accessor :alpha_mod_fix

    # Parse Blip object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [Blip] result of parsing
    def parse(node)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'alphaModFix'
          @alpha_mod_fix = AlphaModFix.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end
