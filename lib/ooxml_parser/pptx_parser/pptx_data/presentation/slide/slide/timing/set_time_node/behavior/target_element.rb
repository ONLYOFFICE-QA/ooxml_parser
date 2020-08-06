# frozen_string_literal: true

module OoxmlParser
  # Class for data for Target Element
  class TargetElement < OOXMLDocumentObject
    attr_accessor :type, :id, :name, :built_in

    def initialize(type = '', id = '', parent: nil)
      @type = type
      @id = id
      super(parent: parent)
    end

    # Parse TargetElement object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [TargetElement] result of parsing
    def parse(node)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'spTgt'
          @type = :shape
          @id = node_child.attribute('spid').value
        end
      end
      self
    end
  end
end
