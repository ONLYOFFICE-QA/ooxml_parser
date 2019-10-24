# frozen_string_literal: true

require_relative 'extension_list/extension'
module OoxmlParser
  # Class for `extLst` data
  class ExtensionList < OOXMLDocumentObject
    attr_accessor :extension_array

    def initialize(parent: nil)
      @extension_array = []
      @parent = parent
    end

    # @return [Array, Extension] accessor
    def [](key)
      @extension_array[key]
    end

    # Parse ExtensionList data
    # @param [Nokogiri::XML:Element] node with ExtensionList data
    # @return [ExtensionList] value of ExtensionList data
    def parse(node)
      node.xpath('*').each do |column_node|
        case column_node.name
        when 'ext'
          @extension_array << Extension.new(parent: self).parse(column_node)
        end
      end
      self
    end
  end
end
