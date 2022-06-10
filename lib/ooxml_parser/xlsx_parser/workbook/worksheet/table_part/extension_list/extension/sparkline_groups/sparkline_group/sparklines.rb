# frozen_string_literal: true

require_relative 'sparklines/sparkline'
module OoxmlParser
  # Class for `sparklines` data
  class Sparklines < OOXMLDocumentObject
    # @return [Array<Sparkline>] list of sparklines
    attr_reader :sparklines

    def initialize(parent: nil)
      @sparklines = []
      super
    end

    # @return [Sparkline] accessor
    def [](key)
      sparklines[key]
    end

    # Parse Sparklines data
    # @param [Nokogiri::XML:Element] node with Sparklines data
    # @return [Sparklines] value of Sparklines data
    def parse(node)
      node.xpath('*').each do |column_node|
        case column_node.name
        when 'sparkline'
          @sparklines << Sparkline.new(parent: self).parse(column_node)
        end
      end
      self
    end
  end
end
