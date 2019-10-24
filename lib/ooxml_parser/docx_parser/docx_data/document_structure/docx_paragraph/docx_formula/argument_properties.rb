# frozen_string_literal: true

require_relative 'argument_properties/argument_size'
module OoxmlParser
  # Class for parsing `m:argPr` object
  class ArgumentProperties < OOXMLDocumentObject
    # @return [ArgumentSize] size of argument
    attr_accessor :argument_size

    # Parse ArgumentProperties
    # @param [Nokogiri::XML:Node] node with ArgumentProperties
    # @return [ArgumentProperties] result of parsing
    def parse(node)
      node.xpath('*').each do |argument_child|
        case argument_child.name
        when 'argSz'
          @argument_size = ArgumentSize.new(parent: self).parse(argument_child)
        end
      end
      self
    end
  end
end
