# frozen_string_literal: true

require_relative 'cell_xfs/xf'
module OoxmlParser
  # Class for parsing `cellXfs`
  class CellXfs < OOXMLDocumentObject
    # @return [Integer] count of cell xfs
    attr_reader :count
    # @return [Array<Xf>] list of xf's
    attr_reader :xf_array

    def initialize(parent: nil)
      @xf_array = []
      @parent = parent
    end

    # Parse CellXfs data
    # @param [Nokogiri::XML:Element] node with CellXfs data
    # @return [CellXfs] value of CellXfs data
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'count'
          @count = value.value.to_i
        end
      end

      node.xpath('*').each do |node_child|
        case node_child.name
        when 'xf'
          @xf_array << Xf.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end
