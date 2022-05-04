# frozen_string_literal: true

require_relative 'matrix/matrix_row'
module OoxmlParser
  # Class for 'm' data
  class Matrix < OOXMLDocumentObject
    attr_accessor :rows

    def initialize(parent: nil)
      @rows = []
      super
    end

    # Parse Matrix object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [Matrix] result of parsing
    def parse(node)
      columns_count = 1
      j = 0
      node.xpath('m:mPr').each do |m_pr|
        m_pr.xpath('m:mcs').each do |mcs|
          mcs.xpath('m:mc').each do |mc|
            mc.xpath('m:mcPr').each do |mc_pr|
              mc_pr.xpath('m:count').each do |count|
                columns_count = count.attribute('val').value.to_i
              end
            end
          end
        end
      end

      node.xpath('m:mr').each do |mr|
        i = 0
        @rows << MatrixRow.new(columns_count, parent: self)
        mr.xpath('m:e').each do |e|
          @rows[j].columns[i] = DocxFormula.new(parent: self).parse(e)
          i += 1
        end
        j += 1
      end

      self
    end
  end
end
