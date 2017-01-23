module OoxmlParser
  # Class for parsing `cond` tags
  class Condition < OOXMLDocumentObject
    attr_accessor :event, :delay, :duration

    # Parse Condition
    # @param node [Nokogiri::XML::Element] node to parse
    # @return [Condition] value of SheetFormatProperties
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'evt'
          @event = value.value
        when 'delay'
          @delay = value.value
        end
      end
      self
    end

    def self.parse_list(conditions_list_node)
      conditions = []
      conditions_list_node.xpath('p:cond').each do |condition_node|
        conditions << Condition.new(parent: conditions).parse(condition_node)
      end
      conditions
    end
  end
end
