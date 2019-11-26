# frozen_string_literal: true

module OoxmlParser
  # Style Parameter Data
  class StyleParametres < OOXMLDocumentObject
    attr_accessor :q_format, :hidden, :name

    def initialize(name = nil,
                   q_format = false,
                   hidden = false,
                   parent: nil)
      @name = name
      @q_format = q_format
      @hidden = hidden
      @parent = parent
    end

    # Parse StyleParametres data
    # @param [Nokogiri::XML:Element] node with StyleParametres data
    # @return [StyleParametres] value of Columns data
    def parse(node)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'name'
          @name = node_child.attribute('val').value
        when 'qFormat'
          @q_format = option_enabled?(node_child)
        end
      end
      self
    end
  end
end
