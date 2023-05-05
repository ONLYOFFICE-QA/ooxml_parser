# frozen_string_literal: true

require_relative 'time_node_list/time_node'
module OoxmlParser
  # Class for parsing TimeNodeList object <p:tnLst>
  class TimeNodeList < OOXMLDocumentObject
    # @return [Array<Object>] list of elements
    attr_reader :elements

    def initialize(parent: nil)
      @elements = []
      super
    end

    # Parse TimeNodeList
    # @param node [Nokogiri::XML::Element] node to parse
    # @return [TimeNodeList] value of TimeNodeList
    def parse(node)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'par'
          @elements << TimeNode.new(:parallel, parent: self).parse(node_child)
        when 'seq'
          @elements << TimeNode.new(:sequence, parent: self).parse(node_child)
        when 'anim'
          @elements << TimeNode.new(:animate, parent: self).parse(node_child)
        when 'set'
          @elements << SetTimeNode.new(parent: self).parse(node_child)
        when 'animEffect'
          @elements << AnimationEffect.new(parent: self).parse(node_child)
        when 'video'
          @elements << :video
        when 'audio'
          @elements << TimeNode.new(:audio, parent: self).parse(node_child)
        end
      end
      self
    end
  end
end
