# frozen_string_literal: true

module OoxmlParser
  class TargetElement < OOXMLDocumentObject
    attr_accessor :type, :id, :name, :built_in

    def initialize(type = '', id = '', parent: nil)
      @type = type
      @id = id
      @parent = parent
    end

    # Parse TargetElement object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [TargetElement] result of parsing
    def parse(node)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'sldTgt'
          @type = :slide
        when 'sndTgt'
          @type = :sound
          @name = node_child.attribute('name').value
          @built_in = node_child.attribute('builtIn').value ? StringHelper.to_bool(node_child.attribute('builtIn').value) : false
        when 'spTgt'
          @type = :shape
          @id = node_child.attribute('spid').value
        when 'inkTgt'
          @type = :ink
          @id = node_child.attribute('spid').value
        end
      end
      self
    end
  end
end
