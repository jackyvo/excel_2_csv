require 'spec_helper'
require_relative '../services/excel_2_csv'

describe 'Excel2CSV' do
	describe '#run' do
  	before do
  		Excel2CSV.new('data/sample-kmp.xlsx').run
  	end

	  it 'exports data to transaction.csv' do
	  	test_file = "spec/fixtures/transactions.csv"
	  	output_file =  "data/output/transactions.csv"

	  	test_data = []
	  	CSV.foreach(test_file, headers: true) do |row|
			  test_data << row.to_hash
			end

			output_data = []
			CSV.foreach(output_file, headers: true) do |row|
			  output_data << row.to_hash
			end

			test_data.each_with_index do |row, index|
				expect(row).to eq(output_data[index])
			end
	  end
	end
end