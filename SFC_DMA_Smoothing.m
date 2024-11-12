%% Read Data
% This section reads data from Excel files for different samples with varying SpiroSi loadings and curing methods (Oven-cured vs. FROMP).
% Each sample represents a combination of SpiroSi loading, curing method, and batch, with data found in the sheet named 'Temperature Ramp - 1'.

Spiro_2_O_1 = readtable('SpiroSi8-2vol%-DCPD-Oven-1.xls', 'Sheet', 'Temperature Ramp - 1');
Spiro_2_O_2 = readtable('SpiroSi8-2vol%-DCPD-Oven-2.xls', 'Sheet', 'Temperature Ramp - 1');
Spiro_2_O_3 = readtable('SpiroSi8-2vol%-DCPD-Oven-3.xls', 'Sheet', 'Temperature Ramp - 1');
Spiro_2_F_1 = readtable('SpiroSi8-2vol%-DCPD-FROMP-1.xls', 'Sheet', 'Temperature Ramp - 1');
Spiro_2_F_2 = readtable('SpiroSi8-2vol%-DCPD-FROMP-2.xls', 'Sheet', 'Temperature Ramp - 1');
Spiro_2_F_3 = readtable('SpiroSi8-2vol%-DCPD-FROMP-3.xls', 'Sheet', 'Temperature Ramp - 1');

Si8_3_O_1 = readtable('Si8-3vol%-DCPD-Oven-1.xls', 'Sheet', 'Temperature Ramp - 1');
Si8_3_O_2 = readtable('Si8-3vol%-DCPD-Oven-2.xls', 'Sheet', 'Temperature Ramp - 1');
Si8_3_O_3 = readtable('Si8-3vol%-DCPD-Oven-3.xls', 'Sheet', 'Temperature Ramp - 1');
Si8_3_F_1 = readtable('Si8-3vol%-DCPD-FROMP-1.xls', 'Sheet', 'Temperature Ramp - 1');
Si8_3_F_2 = readtable('Si8-3vol%-DCPD-FROMP-2.xls', 'Sheet', 'Temperature Ramp - 1');
Si8_3_F_3 = readtable('Si8-3vol%-DCPD-FROMP-3.xls', 'Sheet', 'Temperature Ramp - 1');

Si7_4_O_1 = readtable('Si7-4vol%-DCPD-Oven-1.xls', 'Sheet', 'Temperature Ramp - 1');
Si7_4_O_2 = readtable('Si7-4vol%-DCPD-Oven-2.xls', 'Sheet', 'Temperature Ramp - 1');
Si7_4_O_3 = readtable('Si7-4vol%-DCPD-Oven-3.xls', 'Sheet', 'Temperature Ramp - 1');
Si7_5_F_1 = readtable('Si7-5vol%-DCPD-FROMP-1.xls', 'Sheet', 'Temperature Ramp - 1');
Si7_5_F_2 = readtable('Si7-5vol%-DCPD-FROMP-2.xls', 'Sheet', 'Temperature Ramp - 1');
Si7_5_F_3 = readtable('Si7-5vol%-DCPD-FROMP-3.xls', 'Sheet', 'Temperature Ramp - 1');

DCPD_O_1 = readtable('DCPD-Oven-1.xls', 'Sheet', 'Temperature Ramp - 1');
DCPD_O_2 = readtable('DCPD-Oven-2.xls', 'Sheet', 'Temperature Ramp - 1');
DCPD_O_3 = readtable('DCPD-Oven-1.xls', 'Sheet', 'Temperature Ramp - 1');
DCPD_F_1 = readtable('DCPD-FROMP-1.xls', 'Sheet', 'Temperature Ramp - 1');
DCPD_F_2 = readtable('DCPD-FROMP-2.xls', 'Sheet', 'Temperature Ramp - 1');
DCPD_F_3 = readtable('DCPD-FROMP-1.xls', 'Sheet', 'Temperature Ramp - 1');

%% Smoothing
% Smoothing step to reduce noise in the data using a moving average.
% The specified moving average window size (15) is applied to each sample's data in the process function, which also handles missing values and interpolates data points.

moving_avg_window = 15; % Frame length (window size)

% Smooth the data by processing columns 3 (Temperature) and 6 (tan(δ)).
[Spiro_2_O_3_x_new, Spiro_2_O_3_smoothed] = process(Spiro_2_O_3{:, 3}, Spiro_2_O_3{:, 6}, moving_avg_window);
[Si8_3_O_1_x_new, Si8_3_O_1_smoothed] = process(Si8_3_O_1{:, 3}, Si8_3_O_1{:, 6}, moving_avg_window);
[Si7_4_O_1_x_new, Si7_4_O_1_smoothed] = process(Si7_4_O_1{:, 3}, Si7_4_O_1{:, 6}, moving_avg_window);
[DCPD_O_3_x_new, DCPD_O_3_smoothed] = process(DCPD_O_3{:, 3}, DCPD_O_3{:, 6}, moving_avg_window);

[Spiro_2_F_3_x_new, Spiro_2_F_3_smoothed] = process(Spiro_2_F_3{:, 3}, Spiro_2_F_3{:, 6}, moving_avg_window);
[Si8_3_F_2_x_new, Si8_3_F_2_smoothed] = process(Si8_3_F_2{:, 3}, Si8_3_F_2{:, 6}, moving_avg_window);
[Si7_5_F_3_x_new, Si7_5_F_3_smoothed] = process(Si7_5_F_3{:, 3}, Si7_5_F_3{:, 6}, moving_avg_window);
[DCPD_F_2_x_new, DCPD_F_2_smoothed] = process(DCPD_F_2{:, 3}, DCPD_F_2{:, 6}, moving_avg_window);

%% Subfigures for Oven-Cured Samples
% Plot smoothed tan(δ) versus Temperature for different Oven-cured samples.
% Each sample has a peak temperature annotation.
figure;
plot(Spiro_2_O_3_x_new, Spiro_2_O_3_smoothed, 'Color', [0.6588 0.1333 0.6588], 'LineWidth', 3, 'DisplayName', '2% SpiroSi');
xlabel('Temperature (^{o}C)', 'FontName', 'Arial', 'FontSize', 15);
ylabel('tan(δ)', 'FontName', 'Arial', 'FontSize', 15);
text(170, 1.9, '174.2 ± 3.0', 'Color', [0.6588 0.1333 0.6588], 'FontName', 'Arial', 'FontSize', 12, 'FontWeight', 'Bold');

hold on;
plot(Si8_3_O_1_x_new, Si8_3_O_1_smoothed, 'Color', [0.4824 0.7059 0.9294], 'LineWidth', 3, 'DisplayName', '3% bisSi8');
text(145, 0.85, '169.6 ± 2.1', 'Color', [0.4824 0.7059 0.9294], 'FontName', 'Arial', 'FontSize', 12, 'FontWeight', 'Bold');

hold on;
plot(Si7_4_O_1_x_new, Si7_4_O_1_smoothed, 'Color', [1.0000 0.8314 0.4902], 'LineWidth', 3, 'DisplayName', '4% bisSi7');
text(133, 0.65, '168.5 ± 4.0', 'Color', [1.0000 0.8314 0.4902], 'FontName', 'Arial', 'FontSize', 12, 'FontWeight', 'Bold');

hold on;
plot(DCPD_O_3_x_new, DCPD_O_3_smoothed, 'Color', [0 0 0], 'LineWidth', 3, 'DisplayName', 'DCPD');
text(190, 1.6, '184.7 ± 0.2', 'Color', [0 0 0], 'FontName', 'Arial', 'FontSize', 12, 'FontWeight', 'Bold');

legend('Location', 'Northwest', 'FontName', 'Arial', 'FontSize', 12, 'FontWeight', 'Bold');
xlim([80 215]);
ylim([0 2]);
box off;
ax = gca;
ax.XAxis.LineWidth = 1.2;
ax.YAxis.LineWidth = 1.2;
set(gca, 'TickDir', 'out');
legend boxoff;

%% FROMP Figure
% Plot smoothed tan(δ) versus Temperature for different FROMP samples.
figure;
plot(Spiro_2_F_3_x_new, Spiro_2_F_3_smoothed, 'Color', [0.6588 0.1333 0.6588], 'LineWidth', 3, 'DisplayName', '2% SpiroSi');
xlabel('Temperature (^{o}C)', 'FontName', 'Arial', 'FontSize', 15);
ylabel('tan(δ)', 'FontName', 'Arial', 'FontSize', 15);
text(158, 1.9, '178.9 ± 1.2', 'Color', [0.6588 0.1333 0.6588], 'FontName', 'Arial', 'FontSize', 12, 'FontWeight', 'Bold');

hold on;
plot(Si8_3_F_2_x_new, Si8_3_F_2_smoothed, 'Color', [0.4824 0.7059 0.9294], 'LineWidth', 3, 'DisplayName', '3% bisSi8');
text(150, 1.65, '176.3 ± 1.3', 'Color', [0.4824 0.7059 0.9294], 'FontName', 'Arial', 'FontSize', 12, 'FontWeight', 'Bold');

hold on;
plot(Si7_5_F_3_x_new, Si7_5_F_3_smoothed, 'Color', [1.0000 0.8314 0.4902], 'LineWidth', 3, 'DisplayName', '5% bisSi7');
text(139, 1.45, '166.9 ± 2.0', 'Color', [1.0000 0.8314 0.4902], 'FontName', 'Arial', 'FontSize', 12, 'FontWeight', 'Bold');

hold on;
plot(DCPD_F_2_x_new, DCPD_F_2_smoothed, 'Color', [0 0 0], 'LineWidth', 3, 'DisplayName', 'DCPD');
text(187, 1.95, '184.8 ± 0.4', 'Color', [0 0 0], 'FontName', 'Arial', 'FontSize', 12, 'FontWeight', 'Bold');

legend('Location', 'Northwest', 'FontName', 'Arial', 'FontSize', 12, 'FontWeight', 'Bold');
xlim([80 215]);
ylim([0 2]);
box off;
ax = gca;
ax.XAxis.LineWidth = 1.2;
ax.YAxis.LineWidth = 1.2;
set(gca, 'TickDir', 'out');
legend boxoff;

%% Storage Modulus - Oven-Cured Samples
% Plot smoothed Storage Modulus (E') versus Temperature for different Oven samples, using a logarithmic scale on the y-axis.
figure;
[Spiro_2_O_3_x_new, Spiro_2_O_3_smoothed_7] = process(Spiro_2_O_3{:, 3}, Spiro_2_O_3{:, 7}, moving_avg_window);
semilogy(Spiro_2_O_3_x_new, Spiro_2_O_3_smoothed_7, 'Color', [0.6588 0.1333 0.6588], 'LineWidth', 3, 'DisplayName', '2% SpiroSi');
xlabel('Temperature (^{o}C)', 'FontName', 'Arial', 'FontSize', 15);
ylabel('Storage Modulus (E'', MPa)', 'FontName', 'Arial', 'FontSize', 15);

hold on;
[Si8_3_O_2_x_new, Si8_3_O_2_smoothed_7] = process(Si8_3_O_2{:, 3}, Si8_3_O_2{:, 7}, moving_avg_window);
semilogy(Si8_3_O_2_x_new, Si8_3_O_2_smoothed_7, 'Color', [0.4824 0.7059 0.9294], 'LineWidth', 3, 'DisplayName', '3% bisSi8'); % Smoothed

hold on;
[Si7_4_O_1_x_new, Si7_4_O_1_smoothed_7] = process(Si7_4_O_1{:, 3}, Si7_4_O_1{:, 7}, moving_avg_window);
semilogy(Si7_4_O_1_x_new, Si7_4_O_1_smoothed_7, 'Color', [1.0000 0.8314 0.4902], 'LineWidth', 3, 'DisplayName', '4% bisSi7'); % Smoothed

hold on;
[DCPD_O_2_x_new, DCPD_O_2_smoothed_7] = process(DCPD_O_2{:, 3}, DCPD_O_2{:, 7}, moving_avg_window);
semilogy(DCPD_O_2_x_new, DCPD_O_2_smoothed_7, 'Color', [0 0 0], 'LineWidth', 3, 'DisplayName', 'DCPD'); % Smoothed

legend('Location', 'Southwest', 'FontName', 'Arial', 'FontSize', 12, 'FontWeight', 'Bold');
xlim([80 215]);
ylim([3 3000]);
box off;
ax = gca;
ax.XAxis.LineWidth = 1.2;
ax.YAxis.LineWidth = 1.2;
set(gca, 'TickDir', 'out');
legend boxoff;

%% Storage Modulus - FROMP Samples
% Plot smoothed Storage Modulus (E') versus Temperature for different FROMP samples.
figure;
[Spiro_2_F_3_x_new, Spiro_2_F_3_smoothed_7] = process(Spiro_2_F_3{:, 3}, Spiro_2_F_3{:, 7}, moving_avg_window);
semilogy(Spiro_2_F_3_x_new, Spiro_2_F_3_smoothed_7, 'Color', [0.6588 0.1333 0.6588], 'LineWidth', 3, 'DisplayName', '2% SpiroSi');
xlabel('Temperature (^{o}C)', 'FontName', 'Arial', 'FontSize', 15);
ylabel('Storage Modulus (E'', MPa)', 'FontName', 'Arial', 'FontSize', 15);

hold on;
[Si8_3_F_2_x_new, Si8_3_F_2_smoothed_7] = process(Si8_3_F_2{:, 3}, Si8_3_F_2{:, 7}, moving_avg_window);
semilogy(Si8_3_F_2_x_new, Si8_3_F_2_smoothed_7, 'Color', [0.4824 0.7059 0.9294], 'LineWidth', 3, 'DisplayName', '3% bisSi8'); % Smoothed

hold on;
[Si7_5_F_1_x_new, Si7_5_F_1_smoothed_7] = process(Si7_5_F_1{:, 3}, Si7_5_F_1{:, 7}, moving_avg_window);
semilogy(Si7_5_F_1_x_new, Si7_5_F_1_smoothed_7, 'Color', [1.0000 0.8314 0.4902], 'LineWidth', 3, 'DisplayName', '5% bisSi7'); % Smoothed

hold on;
[DCPD_F_2_x_new, DCPD_F_2_smoothed_7] = process(DCPD_F_2{:, 3}, DCPD_F_2{:, 7}, moving_avg_window);
semilogy(DCPD_F_2_x_new, DCPD_F_2_smoothed_7, 'Color', [0 0 0], 'LineWidth', 3, 'DisplayName', 'DCPD'); % Smoothed

legend('Location', 'Southwest', 'FontName', 'Arial', 'FontSize', 12, 'FontWeight', 'Bold');
xlim([80 215]);
ylim([3 3000]);
box off;
ax = gca;
ax.XAxis.LineWidth = 1.2;
ax.YAxis.LineWidth = 1.2;
set(gca, 'TickDir', 'out');
legend boxoff;

%% Functions
% This function processes data by removing invalid values, averaging duplicate x-values, interpolating, and smoothing.
function [x_new, y_smoothed] = process(x_original, y_original, moving_avg_window)
    valid_indices = isfinite(x_original) & isfinite(y_original);
    x_original = x_original(valid_indices);
    y_original = y_original(valid_indices);

    [x_unique, ~, idx] = unique(x_original);
    y_unique = accumarray(idx, y_original, [], @mean);

    x_new = linspace(min(x_unique), max(x_unique), length(x_unique));
    y_interpolated = interp1(x_unique, y_unique, x_new, 'linear');

    y_smoothed = movmean(y_interpolated, moving_avg_window);
end