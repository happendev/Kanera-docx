import { createHash } from 'node:crypto';
import { describe, expect, it } from 'vitest';
import { ImageBudget } from '../src/image/budget.js';
import { RECOMPRESS_MAX_SCANS } from '../src/image/image.js';
import { decodeJpegPixels, JpegScanLimit, MAX_SCANS } from '../src/image/jpeg-decoder.js';
import { decodePixels, UnreadableJpeg } from '../src/image/pixels.js';
import { convertDocxToImages, docxToPdf } from '../src/node.js';
import { nodeCompression } from '../src/util/node-compression.js';
import { makeDocx } from './helpers/docx.js';
import { inlineImage } from './helpers/sample.js';

// Small Pillow-encoded JPEGs, each with the SHA-1 (first 16 hex digits) of the RGB samples Poppler
// 26.05's pdftoppm draws for it at 1:1, CMYK with the Decode array the PDF writer gives Adobe CMYK.
const cases = [
  {
    name: '4:2:0 chroma at an odd size',
    size: [17, 9],
    rgb: 'fdbdfa07cc82cac9',
    jpeg: [
      '/9j/4AAQSkZJRgABAQAAAQABAAD/2wBDAAYEBQYFBAYGBQYHBwYIChAKCgkJChQODwwQFxQYGBcUFhYaHSUfGhsjHBYWICwg',
      'IyYnKSopGR8tMC0oMCUoKSj/2wBDAQcHBwoIChMKChMoGhYaKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgo',
      'KCgoKCgoKCgoKCgoKCj/wAARCAAJABEDASIAAhEBAxEB/8QAHwAAAQUBAQEBAQEAAAAAAAAAAAECAwQFBgcICQoL/8QAtRAA',
      'AgEDAwIEAwUFBAQAAAF9AQIDAAQRBRIhMUEGE1FhByJxFDKBkaEII0KxwRVS0fAkM2JyggkKFhcYGRolJicoKSo0NTY3ODk6',
      'Q0RFRkdISUpTVFVWV1hZWmNkZWZnaGlqc3R1dnd4eXqDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXG',
      'x8jJytLT1NXW19jZ2uHi4+Tl5ufo6erx8vP09fb3+Pn6/8QAHwEAAwEBAQEBAQEBAQAAAAAAAAECAwQFBgcICQoL/8QAtREA',
      'AgECBAQDBAcFBAQAAQJ3AAECAxEEBSExBhJBUQdhcRMiMoEIFEKRobHBCSMzUvAVYnLRChYkNOEl8RcYGRomJygpKjU2Nzg5',
      'OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6goOEhYaHiImKkpOUlZaXmJmaoqOkpaanqKmqsrO0tba3uLm6wsPE',
      'xcbHyMnK0tPU1dbX2Nna4uPk5ebn6Onq8vP09fb3+Pn6/9oADAMBAAIRAxEAPwDu7zxAmm6d56tGHZ1jTzD8u5jgA8jr0HuR',
      'Vbx54z0xPC93ZW2pW66ndIsS2wdWlCuAWDKM7fk3cn2wckVx3i7/AJAFv/1/2f8A6UJXlFp/yO+tf9dJf/QxXw3C/D2GxuJj',
      'UqfYfN62a0fkdHs4YbApxWrvr6/8Mb9FFFfsZ4R//9k=',
    ].join(''),
  },
  {
    name: 'a progressive 4:2:0 JPEG',
    size: [17, 9],
    rgb: 'fdbdfa07cc82cac9',
    jpeg: [
      '/9j/4AAQSkZJRgABAQAAAQABAAD/2wBDAAYEBQYFBAYGBQYHBwYIChAKCgkJChQODwwQFxQYGBcUFhYaHSUfGhsjHBYWICwg',
      'IyYnKSopGR8tMC0oMCUoKSj/2wBDAQcHBwoIChMKChMoGhYaKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgo',
      'KCgoKCgoKCgoKCgoKCj/wgARCAAJABEDASIAAhEBAxEB/8QAGAAAAgMAAAAAAAAAAAAAAAAAAAUDBAb/xAAXAQADAQAAAAAA',
      'AAAAAAAAAAADBAUG/9oADAMBAAIQAxAAAAF7GmykIl8DZIf/xAAbEAABBAMAAAAAAAAAAAAAAAAEAAIDExQzNP/aAAgBAQAB',
      'BQJ5FcZ5keKi9De1f//EABgRAAIDAAAAAAAAAAAAAAAAAAEDEBIh/9oACAEDAQE/AahaMj//xAAaEQACAgMAAAAAAAAAAAAA',
      'AAACAwEFECFB/9oACAECAQE/AauvW5kEXN4//8QAIRAAAQMBCQAAAAAAAAAAAAAAAgADIRABERITQVFxcrL/2gAIAQEABj8C',
      'vjaUQC4OYUYdaD3D0nubaf/EABwQAQABBQEBAAAAAAAAAAAAAAERABBBUfAhMf/aAAgBAQABPyEIrJQYS0MSMEhg/ZMeTbta',
      'V3N2/9oADAMBAAIAAwAAABCr3//EABkRAAMAAwAAAAAAAAAAAAAAAAERIRBx8f/aAAgBAwEBPxBMQpd3zH//xAAaEQABBQEA',
      'AAAAAAAAAAAAAAABEBEhQWGB/9oACAECAQE/ELePTEQcT//EABwQAQABBAMAAAAAAAAAAAAAAAERABAxYSFR8P/aAAgBAQAB',
      'PxCCxNLIgHkzg2lDuKpIiAnadQyl1Ps9bf/Z',
    ].join(''),
  },
  {
    name: 'Adobe CMYK',
    size: [17, 9],
    rgb: 'df90bdec5cfd1172',
    jpeg: [
      '/9j/7gAOQWRvYmUAZAAAAAAA/9sAQwAGBAUGBQQGBgUGBwcGCAoQCgoJCQoUDg8MEBcUGBgXFBYWGh0lHxobIxwWFiAsICMm',
      'JykqKRkfLTAtKDAlKCko/8AAFAgACQARBEMRAE0RAFkRAEsRAP/EAB8AAAEFAQEBAQEBAAAAAAAAAAABAgMEBQYHCAkKC//E',
      'ALUQAAIBAwMCBAMFBQQEAAABfQECAwAEEQUSITFBBhNRYQcicRQygZGhCCNCscEVUtHwJDNicoIJChYXGBkaJSYnKCkqNDU2',
      'Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6g4SFhoeIiYqSk5SVlpeYmZqio6Slpqeoqaqys7S1tre4ubrC',
      'w8TFxsfIycrS09TV1tfY2drh4uPk5ebn6Onq8fLz9PX29/j5+v/aAA4EQwBNAFkASwAAPwDuvA7QoiPczRwxgqu52CgszBVU',
      'E9yzKAO5OK9m8aeN10LRDdW/lNcPKkMXm/c3Me/IPPIGO5FfMFvApi3uOOlfVNW/i74x0zTPA2raXb6xZprdwq2y2iOry7XI',
      '3hk5Kgx7/mIHUYOSK5z4jfEbSZPCF1ZWWp26apfxJGlruV5grgMwZBnb+73cnjpg5Ir0T4ZeHILnxLp11JEJIoX83rjawXcv',
      'TuDtOP6UV8w14LXvlFe9Wv8AyA4P+wjpv/pfb16N8VP+Rctf+wja/wDo1a+Ro/8AkHr/AL60V4T4r/5Kh4p/6/bj/wBGV4lp',
      'f/I86x/20/8AQ1r274M/8guP/ck/9DFFR109el0V/9k=',
    ].join(''),
  },
  {
    name: 'restart markers',
    size: [64, 48],
    rgb: '9663ada4f5009f7f',
    jpeg: [
      '/9j/4AAQSkZJRgABAQAAAQABAAD/2wBDAAYEBQYFBAYGBQYHBwYIChAKCgkJChQODwwQFxQYGBcUFhYaHSUfGhsjHBYWICwg',
      'IyYnKSopGR8tMC0oMCUoKSj/2wBDAQcHBwoIChMKChMoGhYaKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgo',
      'KCgoKCgoKCgoKCgoKCj/wAARCAAwAEADASIAAhEBAxEB/8QAHwAAAQUBAQEBAQEAAAAAAAAAAAECAwQFBgcICQoL/8QAtRAA',
      'AgEDAwIEAwUFBAQAAAF9AQIDAAQRBRIhMUEGE1FhByJxFDKBkaEII0KxwRVS0fAkM2JyggkKFhcYGRolJicoKSo0NTY3ODk6',
      'Q0RFRkdISUpTVFVWV1hZWmNkZWZnaGlqc3R1dnd4eXqDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXG',
      'x8jJytLT1NXW19jZ2uHi4+Tl5ufo6erx8vP09fb3+Pn6/8QAHwEAAwEBAQEBAQEBAQAAAAAAAAECAwQFBgcICQoL/8QAtREA',
      'AgECBAQDBAcFBAQAAQJ3AAECAxEEBSExBhJBUQdhcRMiMoEIFEKRobHBCSMzUvAVYnLRChYkNOEl8RcYGRomJygpKjU2Nzg5',
      'OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6goOEhYaHiImKkpOUlZaXmJmaoqOkpaanqKmqsrO0tba3uLm6wsPE',
      'xcbHyMnK0tPU1dbX2Nna4uPk5ebn6Onq8vP09fb3+Pn6/90ABAAD/9oADAMBAAIRAxEAPwD0mz1a3hRXmljjQsqAuwALMQqj',
      'nuSQAO5Ires9TtpjKkU0bvE2yRVYEo2AcH0OCD+Ir591XWLi90ZYbSTbP9ptnRtm/YVnjbcR3Axk/SktdZ1KGTTILeLULeW2',
      'mgEzR+a0c7NOPOckfK2UyxZ88NwAQSPzrBcJ/WkpPc9bA5ZOjFJo+hbq7iEbMzKFAySTwBXD6r4gSRytpECufvv3+g/L/CvO',
      'NRuNce3M1rPqD3Ut7fLslldoxH+/MOUJ2hcrFg47gZwQKpQw6g+n3KJPqDI89ooPlyxOo85fNI3uz42nk8LgHHevq8BwngsN',
      'rVhzP8D4LiDHqtVdGhKyTs353s/S349PPuHupn+8+f8AgIrL1TxJZ6TbzyTNHNLFgGFGG4MwJXd/dBAPJ9D1qj4Zjnl8WOk/',
      '237MYrgMk3mFAyyRiMhj8pyu8jbzj7xJBrD13ww+t+Jdbura2a30i0hVC0NrI73Lgylwi/8ALRt+VPI6jrXv08ty5v8AfUoq',
      'K30S+WivqebgeG1iGq1eScPJ6+mn+dz/0PO9T8banfMs/wDaX2WEHCi1k8pAcDPIOT07k98V0vgr4nanotxDFqrHUtNLDfv5',
      'mReclXzyec4bPQAFetVrLwJp66VFNff20bmfeZZba0YeTJhdqiNoyxGON3TK9s8aMPwuLaTZtBeNa3ItU8y3mUOFl28rvGML',
      'njoemea9V1MirQ+rVaMeT/D/AMC69fx1MqdPDU4qEIpJeRxOiXchC8mutt7yUNBErbXmcIGIztyetZegaI5C/Ka9d07wPpUd',
      's0wk+2XW3EUpfaiMCcYC+/XOen4VhgasaOC5qXxy0T7Lq/u0Xm79D9ExefYXLMK5yvzNPlsr62030+/fszFoqrf3YtYgcZdv',
      'ujtWBNcSzn97IzDrjt+VbWPxDh/g7FZzT+sOShT6Nq7fey00Xe61072//9G/XTafdW17ZQ3NjIkltIuUZOmP6Y6Y7V5hDcSw',
      'H91Iyjrjt+VcxoXiPVPDNxNFZyo0W8+ZDIN0bMMjPYj8CM4GelenUyuWPg/Zu0o7Lo+5y0eFKuSyk3UU1K1naz0ve617q2rP',
      'fqK8stvinMsCi60qOSbnc0c5RTzxgFTjj3rS0PxVr3izUzaaLZW9lbhP39zLmUw5z8wPAJ9FIOSD2zjx8Rk+JwtOVbEWhCOr',
      'k2rf5+ml/mVKlKCcpaJHY6B4fVY95ThRk4Fb+m6G9rau8W9eBwpxk9s/nWxo4RbJiAM4xz78VZ+TVoZLIM0AB3K6HlhjHPtz',
      '0r5XD46r7Ok9uWN72vZydk/Pa/yOnB5vNU+SSu5bJ9ban//Sp+JNq63dxR7tkLmIBgMgrwfwzmsytXxXZSad4k1K1m+9HO+D',
      'xypOVP4gg/jWVXoYeypRs76LXvofoeDhCFCEYRSVloiK5nWCPc3JPQetcjqisblpiABIc8DABrc1Ny1zt7KAOv41RljWVCrj',
      'I/lX2GXYONOip/alqRjsJ9aouC33RiV7V8DfI/4R6/2+V9p+1fPjG/ZsXbnvjO/H4+9ePyWUqn5MOPyr2H4Q+DNT08rrN9qU',
      'dnYzw7zar85mXB2lucLjKsCMnBIO3mvmuOsOq+UzoufLJtNX6tO9vn32Ttdo+CzXC1KVJwqrlb28z//Z',
    ].join(''),
  },
  {
    name: 'data cut off mid-MCU',
    size: [64, 48],
    rgb: '86996c3c9ed4812b',
    jpeg: [
      '/9j/4AAQSkZJRgABAQAAAQABAAD/2wBDAAYEBQYFBAYGBQYHBwYIChAKCgkJChQODwwQFxQYGBcUFhYaHSUfGhsjHBYWICwg',
      'IyYnKSopGR8tMC0oMCUoKSj/2wBDAQcHBwoIChMKChMoGhYaKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgo',
      'KCgoKCgoKCgoKCgoKCj/wAARCAAwAEADASIAAhEBAxEB/8QAHwAAAQUBAQEBAQEAAAAAAAAAAAECAwQFBgcICQoL/8QAtRAA',
      'AgEDAwIEAwUFBAQAAAF9AQIDAAQRBRIhMUEGE1FhByJxFDKBkaEII0KxwRVS0fAkM2JyggkKFhcYGRolJicoKSo0NTY3ODk6',
      'Q0RFRkdISUpTVFVWV1hZWmNkZWZnaGlqc3R1dnd4eXqDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXG',
      'x8jJytLT1NXW19jZ2uHi4+Tl5ufo6erx8vP09fb3+Pn6/8QAHwEAAwEBAQEBAQEBAQAAAAAAAAECAwQFBgcICQoL/8QAtREA',
      'AgECBAQDBAcFBAQAAQJ3AAECAxEEBSExBhJBUQdhcRMiMoEIFEKRobHBCSMzUvAVYnLRChYkNOEl8RcYGRomJygpKjU2Nzg5',
      'OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6goOEhYaHiImKkpOUlZaXmJmaoqOkpaanqKmqsrO0tba3uLm6wsPE',
      'xcbHyMnK0tPU1dbX2Nna4uPk5ebn6Onq8vP09fb3+Pn6/9oADAMBAAIRAxEAPwD0mz1a3hRXmljjQsqAuwALMQqjnuSQAO5I',
      'res9TtpjKkU0bvE2yRVYEo2AcH0OCD+Ir591XWLi90ZYbSTbP9ptnRtm/YVnjbcR3Axk/SktdZ1KGTTILeLULeW2mgEzR+a0',
      'c7NOPOckfK2UyxZ88NwAQSPzrBcJ/WkpPc9bA5ZOjFJo+hbq7iEbMzKFAySTwBXD6r4gSRytpECufvv3+g/L/CvONRuNce3M',
      '1rPqD3Ut7fLslldoxH+/MOUJ2hcrFg47gZwQKpQw6g+n3KJPqDI89ooPlyxOo85fNI3uz42nk8LgHHevq8BwngsNrVhzP8D4',
      'LiDHqtVdGhKyTs353s/S349PPuHupn+8+f8AgIrL1TxJZ6TbzyTNHNLFgGFGG4MwJXd/dBAPJ9D1qj4Zjnl8WOk/237MYrgM',
      'k3mFAyyRiMhj8pyu8jbzj7xJBrD13ww+t+Jdbura2a30i0hVC0NrI73Lgylwi/8ALRt+VPI6jrXv08ty5v8AfUoqK30S+Wiv',
      'qebgeG1iGq1eScPJ6+mn+dzE1Pxtqd8yz/2l9lhBwotZPKQHAzyDk9O5PfFdL4K+J2p6LcQxaqx1LTSw37+ZkXnJV88nnOGz',
      '0ABXrVay8CaeulRTX39tG5n3mWW2tGHkyYXaojaMsRjjd0yvbPGjD8Li2k2bQXjWtyLVPMt5lDhZdvK7xjC546HpnmvQdTIq',
      '0Pq1WjHk/wAP/AuvX8dT7SnTw1OKhCKSXkcTol3IQvJrrbe8lDQRK215nCBiM7cnrWXoGiOQvymvXdO8D6VHbNMJPtl1txFK',
      'X2ojAnGAvv1znp+FYYGrGjgual8ctE+y6v7tF5u/Q/RMXn2FyzCucr8zT5bK+ttN9Pv37MxaKq392LWIHGXb7o7/2Q==',
    ].join(''),
  },
  {
    name: 'a missing restart marker',
    size: [64, 48],
    rgb: 'd43a6a85a40269c3',
    jpeg: [
      '/9j/4AAQSkZJRgABAQAAAQABAAD/2wBDAAYEBQYFBAYGBQYHBwYIChAKCgkJChQODwwQFxQYGBcUFhYaHSUfGhsjHBYWICwg',
      'IyYnKSopGR8tMC0oMCUoKSj/2wBDAQcHBwoIChMKChMoGhYaKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgo',
      'KCgoKCgoKCgoKCgoKCj/wAARCAAwAEADASIAAhEBAxEB/8QAHwAAAQUBAQEBAQEAAAAAAAAAAAECAwQFBgcICQoL/8QAtRAA',
      'AgEDAwIEAwUFBAQAAAF9AQIDAAQRBRIhMUEGE1FhByJxFDKBkaEII0KxwRVS0fAkM2JyggkKFhcYGRolJicoKSo0NTY3ODk6',
      'Q0RFRkdISUpTVFVWV1hZWmNkZWZnaGlqc3R1dnd4eXqDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXG',
      'x8jJytLT1NXW19jZ2uHi4+Tl5ufo6erx8vP09fb3+Pn6/8QAHwEAAwEBAQEBAQEBAQAAAAAAAAECAwQFBgcICQoL/8QAtREA',
      'AgECBAQDBAcFBAQAAQJ3AAECAxEEBSExBhJBUQdhcRMiMoEIFEKRobHBCSMzUvAVYnLRChYkNOEl8RcYGRomJygpKjU2Nzg5',
      'OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6goOEhYaHiImKkpOUlZaXmJmaoqOkpaanqKmqsrO0tba3uLm6wsPE',
      'xcbHyMnK0tPU1dbX2Nna4uPk5ebn6Onq8vP09fb3+Pn6/90ABAAD/9oADAMBAAIRAxEAPwD0mz1a3hRXmljjQsqAuwALMQqj',
      'nuSQAO5Ires9TtpjKkU0bvE2yRVYEo2AcH0OCD+Ir591XWLi90ZYbSTbP9ptnRtm/YVnjbcR3Axk/SktdZ1KGTTILeLULeW2',
      'mgEzR+a0c7NOPOckfK2UyxZ88NwAQSPzrBcJ/WkpPc9bA5ZOjFJo+hbq7iEbMzKFAySTwBXD6r4gSRytpECufvv3+g/L/CvO',
      'NRuNce3M1rPqD3Ut7fLslldoxH+/MOUJ2hcrFg47gZwQKpQw6g+n3KJPqDI89ooPlyxOo85fNI3uz42nk8LgHHevq8BwngsN',
      'rVhzP8D4LiDHqtVdGhKyTs353s/S349PPuHupn+8+f8AgIrL1TxJZ6TbzyTNHNLFgGFGG4MwJXd/dBAPJ9D1qj4Zjnl8WOk/',
      '237MYrgMk3mFAyyRiMhj8pyu8jbzj7xJBrD13ww+t+Jdbura2a30i0hVC0NrI73Lgylwi/8ALRt+VPI6jrXv08ty5v8AfUoq',
      'K30S+WivqebgeG1iGq1eScPJ6+mn+dz/0PO9T8banfMs/wDaX2WEHCi1k8pAcDPIOT07k98V0vgr4nanotxDFqrHUtNLDfv5',
      'mReclXzyec4bPQAFetVrLwJp66VFNff20bmfeZZba0YeTJhdqiNoyxGON3TK9s8aMPwuLaTZtBeNa3ItU8y3mUOFl28rvGML',
      'njoemea9V1MirQ+rVaMeT/D/AMC69fx1MqdPDU4qEIpJeRxOiXchC8mutt7yUNBErbXmcIGIztyetZegaI5C/Ka9d07wPpUd',
      's0wk+2XW3EUpfaiMCcYC+/XOen4VhgasaOC5qXxy0T7Lq/u0Xm79D9ExefYXLMK5yvzNPlsr62030+/fszFoqrf3YtYgcZdv',
      'ujtWBNcSzn97IzDrjt+VbWPxDh/g7FZzT+sOShT6Nq7fey00Xe61072/v102n3Vte2UNzYyJJbSLlGTpj+mOmO1eYQ3EsB/d',
      'SMo647flXMaF4j1TwzcTRWcqNFvPmQyDdGzDIz2I/AjOBnpXp1Mrlj4P2btKOy6PuctHhSrkspN1FNStZ2s9L3ute6tqz36i',
      'vLLb4pzLAoutKjkm53NHOUU88YBU44960tD8Va94s1M2mi2VvZW4T9/cy5lMOc/MDwCfRSDkg9s48fEZPicLTlWxFoQjq5Nq',
      '3+fppf5lSpSgnKWiR2OgeH1WPeU4UZOBW/puhva2rvFvXgcKcZPbP51saOEWyYgDOMc+/FWfk1aGSyDNAAdyuh5YYxz7c9K+',
      'Vw+Oq+zpPblje9r2cnZPz2v8jpwebzVPkkruWyfW2p//0qfiTaut3cUe7ZC5iAYDIK8H8M5rMrV8V2UmneJNStZvvRzvg8cq',
      'TlT+IIP41lV6GHsqUbO+i176H6Hg4QhQhGEUlZaIiuZ1gj3NyT0HrXI6orG5aYgASHPAwAa3NTctc7eygDr+NUZY1lQq4yP5',
      'V9hl2DjToqf2pakY7CfWqLgt90Yle1fA3yP+Eev9vlfaftXz4xv2bF2574zvx+PvXj8llKp+TDj8q9h+EPgzU9PK6zfalHZ2',
      'M8O82q/OZlwdpbnC4yrAjJwSDt5r5rjrDqvlM6LnyybTV+rTvb599k7XaPgs1wtSlScKq5W9vM//2Q==',
    ].join(''),
  },
];

const fixture = (name: string) => new Uint8Array(Buffer.from(cases.find((c) => c.name === name)!.jpeg, 'base64'));

/** Where the two-byte `marker` first appears in `d`. */
const find = (d: Uint8Array, marker: number) => d.findIndex((b, i) => b === 0xff && d[i + 1] === marker);

/** How many scans `d` has. */
const scanCount = (d: Uint8Array) => d.filter((b, i) => b === 0xff && d[i + 1] === 0xda).length;

/** A progressive `jpeg` with `n` more DC refinement scans, which carry no data and change nothing. */
function withEmptyScans(jpeg: Uint8Array, n: number): Uint8Array {
  const empty = [0xff, 0xda, 0, 8, 1, jpeg[find(jpeg, 0xc2) + 10], 0, 0, 0, 0x10];
  return new Uint8Array([...jpeg.subarray(0, -2), ...Array.from({ length: n }, () => empty).flat(), 0xff, 0xd9]);
}

/** A progressive `jpeg` that decodes its first scan, the DC coefficients, `n` more times: the same picture. */
function withRepeatedScans(jpeg: Uint8Array, n: number): Uint8Array {
  const start = find(jpeg, 0xda);
  const data = start + 2 + ((jpeg[start + 2] << 8) | jpeg[start + 3]);
  const end = jpeg.findIndex((b, i) => i >= data && b === 0xff && jpeg[i + 1] !== 0 && (jpeg[i + 1] & 0xf8) !== 0xd0);
  const scan = jpeg.subarray(start, end);
  return new Uint8Array([...jpeg.subarray(0, end), ...Array.from({ length: n }, () => [...scan]).flat(), ...jpeg.subarray(end)]);
}

const pictureDocx = (data: Uint8Array) => makeDocx({ body: `<w:p>${inlineImage('rIdImg', 34, 18)}</w:p>`, media: { rIdImg: { name: 'scans.jpg', data } } });

function rgbHash(rgba: Uint8Array): string {
  const rgb = new Uint8Array((rgba.length / 4) * 3);
  for (let i = 0, o = 0; i < rgba.length; i += 4, o += 3) rgb.set(rgba.subarray(i, i + 3), o);
  return createHash('sha1').update(rgb).digest('hex').slice(0, 16);
}

describe('JPEG decoder', () => {
  it.each(cases)('decodes $name exactly as Poppler does', ({ size, rgb, jpeg }) => {
    const img = decodeJpegPixels(new Uint8Array(Buffer.from(jpeg, 'base64')));
    expect([img.width, img.height]).toEqual(size);
    expect(rgbHash(img.data)).toBe(rgb);
  });

  it('rejects a second frame header, whose size the pixel budget never saw', async () => {
    const jpeg = fixture('4:2:0 chroma at an odd size');
    const at = find(jpeg, 0xc0);
    const small = [0xff, 0xc0, 0, 11, 8, 0, 8, 0, 8, 1, 1, 0x11, 0];
    const twoFrames = new Uint8Array([...jpeg.subarray(0, at), ...small, ...jpeg.subarray(at)]);
    expect(() => decodeJpegPixels(twoFrames)).toThrow(/more than one frame/);
    // The page image skips it, so the 8 × 8 frame's reservation can't pay for the 17 × 9 one.
    await expect(decodePixels(twoFrames, nodeCompression, new ImageBudget({ maxImagePixels: 100 }))).rejects.toThrow(UnreadableJpeg);
  });

  it('passes over empty scans and skips a page picture past the scan limit', async () => {
    const { rgb } = cases.find((c) => c.name === 'a progressive 4:2:0 JPEG')!;
    const jpeg = fixture('a progressive 4:2:0 JPEG');
    const scans = scanCount(jpeg);
    expect(rgbHash(decodeJpegPixels(withEmptyScans(jpeg, MAX_SCANS - scans)).data)).toBe(rgb);
    expect(() => decodeJpegPixels(withEmptyScans(jpeg, MAX_SCANS - scans + 1))).toThrow(JpegScanLimit);
    const result = await convertDocxToImages(pictureDocx(withEmptyScans(jpeg, MAX_SCANS)), { dpi: 36 });
    expect(result.warnings).toContain(`Image word/media/scans.jpg is a JPEG with more than ${MAX_SCANS} scans and was skipped`);
  });

  it('recompresses only JPEGs with few enough scans to decode in bounded time', async () => {
    const jpeg = fixture('a progressive 4:2:0 JPEG');
    const scans = scanCount(jpeg);
    const most = withRepeatedScans(jpeg, RECOMPRESS_MAX_SCANS - scans);
    const many = withRepeatedScans(jpeg, RECOMPRESS_MAX_SCANS - scans + 1);
    // The PDF writer recompresses the picture, but embeds it unchanged past the limit.
    const embeds = async (d: Uint8Array) => Buffer.from(await docxToPdf(pictureDocx(d), { imageQuality: 40 })).includes(Buffer.from(d));
    expect([await embeds(most), await embeds(many)]).toEqual([false, true]);
    // Page images decode up to the decoder's own limit.
    const result = await convertDocxToImages(pictureDocx(many), { dpi: 36 });
    expect(result.warnings).toEqual([]);
  });

  it('takes the nearest sample for sampling factors that do not divide the largest', async () => {
    // The second component sampled 3 × 1 against the first's 2 × 2, which libjpeg refuses.
    const original = decodeJpegPixels(fixture('a progressive 4:2:0 JPEG'));
    const fractional = fixture('a progressive 4:2:0 JPEG');
    fractional[find(fractional, 0xc2) + 14] = 0x31;
    const img = decodeJpegPixels(fractional);
    expect([img.width, img.height]).toEqual([original.width, original.height]);
    const result = await convertDocxToImages(pictureDocx(fractional), { dpi: 36 });
    expect(result.warnings).toEqual([]);
  });

  it('rejects arithmetic coding, and a page picture the decoder cannot read is skipped', async () => {
    const sof9 = new Uint8Array([0xff, 0xd8, 0xff, 0xc9, 0, 11, 8, 0, 1, 0, 1, 1, 1, 0x11, 0, 0xff, 0xd9]);
    expect(() => decodeJpegPixels(sof9)).toThrow(/not supported/);
    // The PDF embeds it undecoded, so the page image still renders.
    const result = await convertDocxToImages(pictureDocx(sof9), { dpi: 36 });
    expect(result.images).toHaveLength(1);
    expect(result.warnings).toContain('Image word/media/scans.jpg is a JPEG that could not be decoded and was skipped');
  });
});
