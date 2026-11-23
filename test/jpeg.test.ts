import { createHash } from 'node:crypto';
import { describe, expect, it } from 'vitest';
import { decodeJpegPixels } from '../src/image/jpeg-decoder.js';

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
];

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

  it('rejects arithmetic coding, leaving the picture to the fallback decoder', () => {
    const sof9 = new Uint8Array([0xff, 0xd8, 0xff, 0xc9, 0, 11, 8, 0, 1, 0, 1, 1, 1, 0x11, 0, 0xff, 0xd9]);
    expect(() => decodeJpegPixels(sof9)).toThrow(/not supported/);
  });
});
