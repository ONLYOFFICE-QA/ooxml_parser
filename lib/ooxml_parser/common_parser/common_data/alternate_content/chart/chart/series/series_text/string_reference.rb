require_relative 'string_reference/string_cache'
module OoxmlParser
  # Class for parsing `c:strRef` object
  class StringReference < OOXMLDocumentObject
    # @return [String] formula
    attr_accessor :formula
    # @return [StringCache] cache of string
    attr_accessor :cache

    # Parse Order
    # @param [Nokogiri::XML:Node] node with Order
    # @return [Order] result of parsing
    def self.parse(node)
      str_ref = StringReference.new
      node.xpath('*').each do |reference_child|
        case reference_child.name
        when 'f'
          str_ref.formula = reference_child.text
        when 'strCache'
          str_ref.cache = StringCache.parse(reference_child)
        end
      end
      str_ref
    end
  end
end
