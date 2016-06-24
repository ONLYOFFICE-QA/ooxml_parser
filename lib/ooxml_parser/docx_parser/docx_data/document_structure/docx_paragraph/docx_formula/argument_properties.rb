require_relative 'argument_properties/argument_size'
module OoxmlParser
  # Class for parsing `m:argPr` object
  class ArgumentProperties
    # @return [ArgumentSize] size of argument
    attr_accessor :argument_size

    # Parse ArgumentProperties
    # @param [Nokogiri::XML:Node] node with ArgumentProperties
    # @return [ArgumentProperties] result of parsing
    def self.parse(node)
      argument_propertie = ArgumentProperties.new
      node.xpath('*').each do |argument_child|
        case argument_child.name
        when 'argSz'
          argument_propertie.argument_size = ArgumentSize.parse(argument_child)
        end
      end
      argument_propertie
    end
  end
end
