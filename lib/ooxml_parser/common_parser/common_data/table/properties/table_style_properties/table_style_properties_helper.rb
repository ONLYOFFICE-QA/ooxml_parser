# frozen_string_literal: true

module OoxmlParser
  # Helper method for working with TableStyleProperties
  module TableStylePropertiesHelper
    # [Hash] short names for table styles
    TABLE_STYLES_NAMES_HASH = { first_column: :firstCol,
                                last_column: :lastCol,
                                whole_table: :wholeTbl,
                                banding_1_horizontal: :band1Horz,
                                banding_2_horizontal: :band2Horz,
                                banding_1_vertical: :band1Vert,
                                banding_2_vertical: :band2Vert,
                                southeast_cell: :seCell,
                                southwest_cell: :swCell,
                                northeast_cell: :neCell,
                                northwest_cell: :nwCell,
                                first_row: :firstRow,
                                last_row: :lastRow }.freeze

    TABLE_STYLES_NAMES_HASH.each do |key, value|
      define_method(key) do
        found = nil
        @table_style_properties_list.each do |table_style|
          # Cannot just use return in block
          # because of TruffleRuby bug (cause LocalJumpError)
          # https://github.com/oracle/truffleruby/issues/2438
          if table_style.type == value
            found = table_style
            break
          end
        end
        found
      end
    end

    # Fill all empty tables styles with default value
    # To make last changes in parsing table styles compatible
    # with `ooxml_parser` 0.1.2 and earlier
    # @return [Nothing]
    def fill_empty_table_styles
      TABLE_STYLES_NAMES_HASH.each_value do |current_table_style|
        style_exist = false
        @table_style_properties_list.each do |existing_style|
          style_exist = true if existing_style.type == current_table_style
        end
        next if style_exist

        @table_style_properties_list << TableStyleProperties.new(type: current_table_style)
      end
    end
  end
end
